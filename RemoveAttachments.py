#!/opt/bin/python
# -*- coding: utf-8 -*-
# Author: Daniel Drake, November 2009
# Copyright HUDORA GmbH 2009

import sys
import logging
import re
import email.parser
import email.utils
import couchdb.client
import socket
import urllib
import hashlib

from datetime import date
from optparse import OptionParser
from imaplib import IMAP4, IMAP4_SSL
from types import ListType, TupleType

missing = object()

excluded_uids = ["339"]


def get_filename_from_part(part):
    """Get filename from a message part.

    This is an extension on Python's email.message.get_filename().  Python's function does not look in
    the Content-Type header for the filename, whereas I have a few emails that have the filename stuffed
    there.
    """
    filename = part.get_filename(missing)
    if filename is missing:
        filename = part.get_param('filename', missing, 'content-type')
        if filename is not missing:
            filename = email.utils.collapse_rfc2231_value(filename).strip()
    
    filename = str(filename).replace('\n','')
    filename = re.sub('[^\x21-\x7E]*', '', filename)
    return filename


def split_response(resp):
    """Itemize an IMAP response into its elements.

    Returns a list of tuples, where each tuple corresponds to an element. The first item is the type
    ("plist" or "string") and the second parameter is the value.
    Only supports single-dimension parameter lists and strings.

    >>> split_response('"INBOX"')
    [('string', 'INBOX')]
    >>> split_response('"INBOX" "." (flag1 flag2)')
    [('string', 'INBOX'), ('string', '.'), ('plist', 'flag1 flag2')]
    """
    ret = []

    while True:
        resp = resp.lstrip()
        if len(resp) == 0:
            break
        if resp[0] == "(":
            endpos = resp.find(")")
            ret.append(("plist", resp[1:endpos]))
            resp = resp[endpos + 1:]
        elif resp[0] == '"':
            resp = resp[1:]
            endpos = resp.find('"')
            ret.append(("string", resp[:endpos]))
            resp = resp[endpos + 1:]
    return ret


def parse_flags(flagstr):
    """Retrieve IMAP mail flags from a string.

    Parses the "FLAGS (Flag1 Flag2 Flag3)" section from a string, returning a list of space-separated flags
    in a string.

    >>> parse_flags("FLAGS (Flag1 Flag2 Flag3)")
    'Flag1 Flag2 Flag3'
    >>> parse_flags("(UID 321 FLAGS (Foo Bar))")
    'Foo Bar'
    """
    if (isinstance(flagstr, ListType) or isinstance(flagstr, TupleType)):
        flagstr = flagstr[0]
    m = re.match(r".*FLAGS\s+\(([^)]+)\).*", flagstr)
    if not m:
        return None
    return m.groups()[0]


def parse_uid(uidstr):
    """Retrieve IMAP mail flags from a string.

    Parses the "UID 123" section from a string, returning the UID as a string.

    >>> parse_uid("UID 1234")
    '1234'
    >>> parse_uid("(UID 321 FLAGS (Foo Bar))")
    '321'
    """
    m = re.match(r".*UID\s+(\d+).*", uidstr)
    if not m:
        return None
    return m.groups()[0]


def parse_internaldate(datestr):
    """Retrieve IMAP internal date from a string.

    Parses the "INTERNALDATE 14-Nov-2009 16:05:24 +0000" section from a string, returning the date as a
    string.

    >>> parse_internaldate('(UID 321 INTERNALDATE "14-Nov-2009 16:05:24 +0000" FLAGS (Foo))')
    '14-Nov-2009 16:05:24 +0000'
    """
    if (isinstance(datestr, ListType) or isinstance(datestr, TupleType)):
        datestr = datestr[0]
    m = re.match(r".*INTERNALDATE\s+\"([^\"]+)\".*", datestr)
    if not m:
        return None
    return m.groups()[0]


class RemoveAttachmentsException(Exception):
    """Exception type generated by the RemoveAttachments class."""
    pass


class RemoveAttachments(object):
    """Remove attachments program class.

    This program goes through all your email looking for attachments. It can either archive the attachments
    in a CouchDB database, delete them from the original email (leaving an explanatory message inside the
    email), or both.
    """

    def __init__(self, server, port, ssl, username, password, only_mailbox=None, cdb_server=None,
                 cdb_db=None, remove=False, eat_more_attachments=False, gmail=False, min_size=0,
                 before_date=None):
        """Constructor.

        Arguments:
        server -- IMAP server to connect to (string)
        port -- IMAP port on server (int)
        ssl -- whether to use SSL (bool)
        username -- IMAP username (string)
        password -- IMAP password (string)
        only_mailbox -- Only work in this mailbox, non-recursively (default: recurse through all mailboxes)
        cdb_server -- CouchDB URL (string) if you want archiving, None if you don't
        cdb_db -- CouchDB database name to use (string)
        remove -- if True, attachments are deleted from mails
        eat_more_attachments -- looser criteria for detecting attachments
        gmail -- Enable Gmail quirks
        min_size -- minimum size of mails to examine, in kB (int, 0 to disable)
        before_date -- only look at mails that arrived before this date (datetime.date or None to disable)
        """
        if None in (server, username, password):
            raise RemoveAttachmentsException("Server, username and password are all required.")

        if not remove and cdb_server is None:
            raise RemoveAttachmentsException("No action specified (expected a CouchDB server, or the " \
                                             "remove option, or both)")

        logging.debug("Connecting to IMAP server")
        try:
            if ssl:
                self.imap = IMAP4_SSL(server, port)
            else:
                self.imap = IMAP4(server, port)
        except socket.error, e:
            raise RemoveAttachmentsException("Could not connect to IMAP server: " + str(e))
        except Exception, e:
            logging.exception(e)
            raise RemoveAttachmentsException("Could not connect to IMAP server")

        try:
            self.imap.login(username, password)
        except IMAP4.error, e:
            raise RemoveAttachmentsException("Could not authenticate: " + str(e))
        except Exception, e:
            logging.exception(e)
            raise RemoveAttachmentsException("Could not authenticate")

        self.searchstr = ''
        self.remove = remove
        self.min_size = min_size * 1024
        self.before_date = before_date
        self.eat_more_attachments = eat_more_attachments or False
        self.only_mailbox = only_mailbox
        self.gmail = gmail or False
        self.cdb_server = cdb_server
        self.cdb_db = cdb_db
        
        if cdb_server:
            cdb_db = cdb_db or "attachments"
            logging.debug("Connecting to CouchDB")
            try:
                self.db_server = couchdb.client.Server(cdb_server)
                if cdb_db in self.db_server:
                    self.db = self.db_server[cdb_db]
                else:
                    self.db = self.db_server.create(cdb_db)
            except socket.error, e:
                raise RemoveAttachmentsException("CouchDB socket error: " + str(e))
            except Exception, e:
                logging.exception(e)
                raise RemoveAttachmentsException("CouchDB error")
        else:
            self.db = None

    def run(self):
        """Run the filtering process"""
        # Construct IMAP search string
        if self.before_date is not None:
            self.searchstr += ' BEFORE %s' % self.before_date.strftime("%d-%b-%Y")
        if self.min_size > 0:
            self.searchstr += ' LARGER %d' % self.min_size
        if self.searchstr == '':
            self.searchstr = 'ALL'
        self.searchstr = '(' + self.searchstr.strip() + ')'
        logging.debug("IMAP search query is %s", self.searchstr)

        if self.only_mailbox is not None:
            logging.debug("Limited to only mailbox %s", self.only_mailbox)
            self._process_mailbox(self.only_mailbox)
        else:
            logging.debug("Processing all mailboxes")
            typ, data = self.imap.list('')
            if typ != "OK":
                raise RemoveAttachmentsException("LIST not OK")
            for ent in data:
                split = split_response(ent)
                if len(split) != 3:
                    logging.error("Unrecognised LIST response entry: %s", str(split))
                    continue
                if split[2][0] != "string":
                    logging.error("Unrecognised non-string response %s", split[2][1])
                    continue
                self._process_mailbox(split[2][1])

        self.imap.logout()

    def _lookup_uids(self, mset):
        logging.debug("UID lookup...")
        uids = []
        for num in mset.split():
            typ, msg = self.imap.fetch(num, '(UID)')
            if typ != 'OK':
                logging.warning("FETCH UID not OK, skipping message %s", num)
            else:
                uid_nr = parse_uid(msg[0])
                if uid_nr not in excluded_uids:
                    uids.append(uid_nr)
                else:
                    logging.info("Skipping excluded UID %s", uid_nr)

        return uids

    def _process_mailbox(self, mailbox):
        try:
            self.__process_mailbox(mailbox)
        except IMAP4.readonly, e:
            logging.info("Skipping mailbox %s as it is read-only", mailbox)
        except Exception, e:
            logging.warning("Error processing mailbox %s", mailbox)
            logging.exception(e)

    def __process_mailbox(self, mailbox):
        """Process an individual IMAP mailbox.

        If removing is enabled, the mailbox will be expunged during this function call.
        """
        logging.debug("Processing mailbox %s", mailbox)
        self.imap.select(mailbox)

        # recurse through every mailbox by looking at imap.list() output
        typ, data = self.imap.search(None, self.searchstr)
        if typ != "OK":
            raise Exception("Search not OK")
        if len(data) == 0 or len(data[0]) == 0:
            return

        # In theory, we should be able to operate directly on the message sequence numbers returned by
        # FETCH. According to the IMAP4rev1 specs, none of the operations we perform in the inner loop
        # will affect the message sequencing. All new mails that we create (the ones without attachments)
        # are guaranteed to have higher sequence numbers and UIDs than the old ones.
        # Well, Gmail's IMAP server doesn't follow the specs here. When you APPEND a new mail to the
        # mailbox, it often ends up having a sequence number within the range that you had previously
        # SEARCHed for.
        # We work around this by looking up the UIDs for all messages and working entirely with UIDs instead
        # of sequence numbers.
        uids = self._lookup_uids(data[0])

        for uid in uids:
            logging.debug("Retrieve mail with uid %s", uid)
            typ, msg = self.imap.uid('FETCH', uid, '(FLAGS INTERNALDATE BODY.PEEK[])')
            if typ != "OK":
                raise Exception("FETCH not OK")
            if len(msg) < 1 or len(msg[0]) < 2:
                raise Exception("Malformed FETCH response")
            try:
                flags = None
                for item in msg:
                    flags = parse_flags(item)
                    if flags is not None:
                        break

                idate = None
                for item in msg:
                    idate = parse_internaldate(item)
                    if idate is not None:
                        break
                self._process_mail(mailbox, uid, flags, idate, msg[0][1])
            except Exception, e:
                logging.warning("Error processing mail %s", uid)
                logging.exception(e)

        if self.remove:
            self.imap.expunge()
        self.imap.close()

    def _part_is_attachment(self, part):
        """Determine whether a mail part is an attachment by looking at its headers"""
        attachment = part.get_param('attachment', missing, 'content-disposition')
        if attachment is not missing:
            return True

        # hmm, some of my mails have PDFs attached which aren't marked as attachments. so let's add this
        # optional 2nd metric for attachment detection
        if self.eat_more_attachments \
                and 'Content-Transfer-Encoding' in part \
                and part['Content-Transfer-Encoding'] == 'base64' \
                and get_filename_from_part(part) is not None:
            return True

        return False

    def _process_mail(self, mailbox, uid, flags, idate, msg):
        """Process the attachments (if any) on an individual mail"""
        parser = email.parser.Parser()
        mail = parser.parsestr(msg)
        found_attachment = False
        doc_id = None

        if 'message-id' not in mail:
            mail['message-id'] = "%s@fakeid.hudora.biz" % hashlib.sha1(repr(mail._headers)).hexdigest()
            logging.warning(" mail %s: no Message-ID, using fake-id %s", uid, mail['message-id'])

        logging.debug("Message-ID: %s", mail['message-id'])

        # quick first pass to see if we have an attachment
        for part in mail.walk():
            if self._part_is_attachment(part):
                found_attachment = True
                break

        if not found_attachment:
            logging.debug("No attachments --> skip (%d bytes)" % len(str(mail)))
            return

        if self.db is not None:
            doc_id = self._save_mail_to_db(mailbox, mail)
        if self.remove:
            self._remove_attachments(mail, doc_id, mailbox, uid, flags, idate)

    def _remove_attachments(self, mail, doc_id, mailbox, uid, flags, idate):
        """Remove the attachments from a mail, replacing them with explanatory messages."""
        logging.debug("Remove attachments")

        if idate is not None:
            idate = '"' + idate + '"'

        modified = False
        for part in mail.walk():
            if not self._part_is_attachment(part):
                continue

            notice = "An attachment in this email was moved into the attachments database, or removed.\n"
            notice += "Filename: %s\n" % get_filename_from_part(part)
            notice += "Content type: %s\n" % part.get_content_type()
            if doc_id:
                notice += "Database document ID: %s\n" % doc_id
                notice += "http://intern.hudora.biz/attachmentarchive/" # %s\n" % (urllib.quote(doc_id))
            part.set_payload(notice)
            for k, v in part.get_params():
                part.del_param(k)
            part.set_type('text/plain')
            del part['Content-Transfer-Encoding']
            del part['Content-Disposition']
            del part['Content-Type']
            modified = True

        if modified:
            self.imap.append(mailbox, flags, idate, mail.as_string())
            if self.gmail:
                # Deleting a mail from a Gmail IMAP server only deletes its labels. To delete it properly,
                # we have to move it into Trash. And moving an email in IMAP terms is making a copy,
                # marking the original as deleted, then expunging. The expunging is done later. Gmail will
                # then delete the copy of the email with the attachments in 30 days time.
                self.imap.uid('COPY', uid, "[Gmail]/Trash")

            # unfortunately, with gmail, this STORE command will add a \Seen tag to the mail we just
            # appended. In theory we could avoid this by putting the following STORE command above the
            # COPY command, but in practice this actually overrides the COPY behaviour such that the mail
            # is not moved into the Trash, so we have no option... :(
            self.imap.uid('STORE', uid, '+FLAGS', '(\\Deleted)')

    def _save_mail_to_db(self, mailbox, mail):
        """Save the attachments from a mail in a CouchDB document.

        Detects if the mail is already there (based on mailbox and message ID) - will not create duplicates.
        """
        doc_id = mailbox + "@@" + re.sub('[^\x21-\x7E]*', '', mail['message-id'])
        doc = {"_attachments": {}}
        if doc_id in self.db:
            if self.db[doc_id]['done'] == True:
                logging.debug("Message is already in CouchDB")
                # document already in db
                return doc_id
            else:
                # we were aborted somewhere after creating the document
                # and finishing uploading the attachments. so delte the old
                # attempt and start again.
                logging.warning("Incomplete upload detected, retrying...")
                doc = self.db[doc_id]
        else:
            logging.debug("Add mail to CouchDB")
        save_hdrs = ('to', 'from', 'date', 'subject', 'in-reply-to', 'references', 'message-id')
        for hdr in save_hdrs:
            if hdr in mail:
                doc[hdr] = mail[hdr]
        doc['mailbox'] = mailbox
        doc['_id'] = doc_id
        doc['done'] = False
        new_doc_id = self.db.create(doc)

        for part in mail.walk():
            attachment = part.get_param('attachment', missing, 'content-disposition')
            if attachment is missing:
                continue

            payload = part.get_payload(decode=True)
            if not payload:
                continue

            filename = get_filename_from_part(part)
            params = part.get_params()
            if len(params) > 0 and '/' in params[0][0]:
                mimetype = params[0][0]
            else:
                mimetype = 'application/octet-stream'

            print len(payload)
            print repr(filename)
            print repr(mimetype)
            print new_doc_id
            self.db.put_attachment(self.db[new_doc_id], payload, filename, mimetype)
            logging.debug("Added attachment %s", filename)

        # modify document to mark it as complete
        doc = self.db[new_doc_id]
        doc['done'] = True
        self.db[new_doc_id] = doc
        logging.debug("Created document %s", new_doc_id)
        return new_doc_id


def die(msg):
    """Abort with an error message."""
    logging.error(msg)
    sys.exit(1)


def parse_date(datestr):
    """Parse a string in the form YYYY-MM-DD into a datetime.date object.

    Aborts if the date was invalid.
    >>> parse_date("2009-11-14")
    datetime.date(2009, 11, 14)
    >>> parse_date("1-2-3")
    datetime.date(1, 2, 3)
    """
    parts = datestr.split("-")
    if len(parts) != 3:
        die("Date format must be YYYY-MM-DD")

    try:
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except:
        die("Date parsing error. Use format YYYY-MM-DD")


def main():
    """Parse command-line arguments and invoke the RemoveAttachments program class."""
    parser = OptionParser()
    parser.add_option("-s", "--server", help="IMAP4 server")
    parser.add_option("--port", help="IMAP4 server port")
    parser.add_option("--ssl", help="Use SSL connectivity", action="store_true")
    parser.add_option("-u", "--username", help="IMAP4 username")
    parser.add_option("-p", "--password", help="IMAP4 password")
    parser.add_option("--only-mailbox", help="Specify a single mailbox to process (default: all, recursive)")
    parser.add_option("--couchdb-server", help="CouchDB server address")
    parser.add_option("--couchdb-db", help="CouchDB database name (default: attachments)")
    parser.add_option("--before-date",
                      help="Only check messages sent before this date (format YYYY-MM-DD)")
    parser.add_option("--min-size", default=5000,
                      help="Ignore mails smaller than this size, in kB [%default]")
    parser.add_option("-r", "--remove", help="Remove attachments after processing", action="store_true")
    parser.add_option("--eat-more-attachments", help="Use looser criteria for detecting attachments",
                      action="store_true")
    parser.add_option("--gmail", help="Enable Gmail quirks mode (see README) (default off)",
                      action="store_true")
    parser.add_option("-v", "--verbose", help="Log debug messages", action="store_true")
    options = parser.parse_args()[0]

    if options.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    if options.before_date is not None:
        before_date = parse_date(options.before_date)
    else:
        before_date = None

    if options.min_size is not None:
        try:
            min_size = int(options.min_size)
        except ValueError:
            die("--min-size requires integer argument")
    else:
        min_size = 0

    if options.port is not None:
        try:
            port = int(options.port)
        except ValueError:
            die("--port requires integer argument")
    else:
        if options.ssl:
            port = 993
        else:
            port = 143

    try:
        RemoveAttachments(options.server, port, options.ssl, options.username, options.password,
                          options.only_mailbox, options.couchdb_server, options.couchdb_db, options.remove,
                          options.eat_more_attachments, options.gmail, min_size, before_date).run()
    except RemoveAttachmentsException, e:
        logging.error(e)
        sys.exit(1)


def doctests():
    """Run doctests"""
    import doctest
    doctest.testmod()


if __name__ == "__main__":
    main()

