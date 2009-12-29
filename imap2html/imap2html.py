#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import email
import mimetypes
import cgi
import socket
import time
import logging
import re
import urllib

from email.header import decode_header
from imaplib import IMAP4
from imaplib import IMAP4_SSL
from optparse import OptionParser

options = None

message_template = """
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>{title}</title>
        <script language="JavaScript" type="text/javascript">
            function show(id)
            {{
                var item = document.getElementById(id);
                item.style.display = (item.style.display == "none" ? "block" : "none");
            }}
        </script>
    </head>
    <body>
        <div id="base-headers" style="display: block;">
            {base_headers}
            <form>
                <input type="button" value="show all headers" onClick="show('base-headers'), show('all-headers')">
            </form>
        </div>
        <div id="all-headers" style="display: none;">
            {all_headers}
            <form>
                <input type="button" value="hide headers" onClick="show('base-headers'), show('all-headers')">
            </form>
        </div>
        <p>
            {text}
        </p>
        <p>
            {html}
        </p>
        <p>
            {files}
        </p>
    <body>
<html>"""

list_template = """
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>{title}</title>
    </head>
    <body>
        <p>
            <b>Show</b>&nbsp;
            <a href="index.html">all</a>&nbsp;|&nbsp;
            <a href="by-sender.html">by sender</a>&nbsp;|&nbsp;
            <a href="by-subject.html">by subject</a>&nbsp;|&nbsp;
            <a href="by-date.html">by date</a>
        </p>
        <p>
            <table width="100%">
                <tr align="left">
                    <th>From</th>
                    <th>Subject</th>
                    <th>Date and Time</th>
                </tr>
                {table}
            </table>
        </p>
    <body>
<html>"""

overview_template = """
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <title>{title}</title>
    </head>
    <body>
        <div align="center">
            {body}
        </div>
    <body>
<html>"""


def is_attachment(part):
    if part['content-disposition'] and part['content-disposition'].startswith('attachment'):
        return True
    else:
        return False


def extract_header(key, headers):
    for item in headers:
        if item[0] == key:
            return item[1]
    for item in headers:
        if item[0] == key.lower():
            return item[1]
    return ''


def decode_string(string):
    result = ''
    try:
        for text, enc in decode_header(string):
            if enc:
                result += unicode(text, enc, 'ignore').encode('utf-8', 'ignore')
            else:
                result += text
    except UnicodeEncodeError:
        result = str(re.sub('[^\x21-\x7E]*', '', result))
    return result


def save_file(filename, data, *dirs):
    path = options.outputdir
    for subdir in dirs:
        path = os.path.join(path, subdir)
        if not os.path.isdir(path):
            os.mkdir(path)
    path = os.path.join(path, filename)

    fd = open(path, 'wb')
    fd.write(data)
    fd.close()


def process_options():
    """Process options passed via command line args."""
    global options
    parser = OptionParser()
    parser.add_option("--server", dest="server", type="string",
                    help="IMAP hostname")
    parser.add_option("--user", dest="user", type="string",
                    help="IMAP user")
    parser.add_option("--password", dest="password", type="string",
                    help="IMAP password")
    parser.add_option("--ssl", dest="ssl", action="store_true", default=False,
                    help="use SSL")
    parser.add_option("--outputdir", dest="outputdir", type="string",
                    default="./mailarchive", help="output directory [%default]")
    parser.add_option("--mindate", dest="mindate", type="string",
                    help="only check messages send before this date (format YYYY-MM-DD)")
    parser.add_option("--minsize", dest="minsize", type="int", default=0,
                    help="only check messages bigger than this size, in kB [%default]")
    parser.add_option("--remove", dest="remove", action="store_true", default=False,
                    help="remove messages from server after processing")
    parser.add_option("--debug", dest="debug", action="store_true", default=False,
                    help="log debug messages")

    options, args = parser.parse_args()

    if args:
        parser.print_help()
        sys.exit(0)

    if not options.server:
        print 'Server option requred, exit'
        sys.exit(1)
    if not options.user:
        print 'User option required, exit'
        sys.exit(1)
    if not options.password:
        print 'Password option required, exit'
        sys.exit(1)


def parse_message(data):
    "Parse raw mail into components"
    text = []
    html = []
    headers = []
    attachments = []
    message = email.message_from_string(data)

    for item in message.items():
        headers.append((item[0], decode_string(item[1])))

    for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        elif part.get_content_type() == 'text/plain' and not is_attachment(part):
            payload = part.get_payload(decode=True)
            enc = part.get_content_charset(None)
            if enc and enc != 'utf-8':
                payload = unicode(payload, enc, 'replace').encode('utf-8', 'replace')
            text.append(payload)
        elif part.get_content_type() == 'text/html' and not is_attachment(part):
            html.append(part.get_payload(decode=True))
        else:
            filename = part.get_filename()
            if filename:
                filename = decode_string(filename)
                filename = os.path.basename(filename)
            else:
                ext = mimetypes.guess_extension(part.get_content_type())
                if not ext:
                    ext = '.bin'
                filename = 'file' + ext
            data = part.get_payload(decode=True)
            if data:
                attachments.append((filename, data))

    return ((text, html), headers, attachments)


def process_message(uid, data):
    message, headers, attachments = parse_message(data)

    html = generate_message(message, headers, attachments)
    save_file('message.html', html, str(uid))

    counter = 0
    for item in message[1]:
        save_file('original-html-%d.html' % counter, item, str(uid))
        counter += 1

    counter = 0
    for item in attachments:
        save_file(item[0], item[1], str(uid), 'part-' + str(counter))
        counter += 1

    return headers


def process_overviews(msg_list):
    file_id = 0
    template = """
    <a href="{filename}.html">{linkname}</a>({count})<br>
    """

    by_sender = {}
    by_subject = {}
    by_date = {}

    for item in msg_list:
        by_sender[item[1]] = by_sender.get(item[1], 0) + 1
        by_subject[item[2]] = by_subject.get(item[2], 0) + 1

        strtime = time.strftime('%Y-%m-%d', item[3])
        by_date[strtime] = by_date.get(strtime, 0) + 1

    links = ''
    for key, val in by_sender.iteritems():
        sublist = filter(lambda x: x[1] == key, msg_list)
        html = generate_list_of_messages(sublist, key)
        save_file(str(file_id) + '.html', html)

        links = links + template.format(filename=str(file_id),
                                        linkname=cgi.escape(key),
                                        count=val)

        file_id = file_id + 1

    html = overview_template.format(body=links, title='Messages by sender')
    save_file('by-sender.html', html)

    links = ''
    for key, val in by_subject.iteritems():
        sublist = filter(lambda x: x[2] == key, msg_list)
        html = generate_list_of_messages(sublist, key)
        save_file(str(file_id) + '.html', html)

        links = links + template.format(filename=str(file_id),
                                        linkname=cgi.escape(key),
                                        count=val)

        file_id = file_id + 1

    html = overview_template.format(body=links, title='Messages by subject')
    save_file('by-subject.html', html)

    links = ''
    for key, val in by_date.iteritems():
        sublist = filter(lambda x: time.strftime('%Y-%m-%d', x[3]) == key, msg_list)
        html = generate_list_of_messages(sublist, key)
        save_file(str(file_id) + '.html', html)

        links = links + template.format(filename=str(file_id),
                                        linkname=cgi.escape(key),
                                        count=val)

        file_id = file_id + 1

    html = overview_template.format(body=links, title='Messages by date')
    save_file('by-date.html', html)

    html = generate_list_of_messages(msg_list, 'All messages')
    save_file('index.html', html)


def generate_message(message, headers, attachments):
    """Generate a html representation of message."""
    subject_header = extract_header('Subject', headers)
    from_header = extract_header('From', headers)
    date_header = extract_header('Date', headers)

    base_headers = '<b>From: </b>%s<br>' % cgi.escape(from_header)
    if subject_header:
        title = subject_header
        base_headers += '<b>Subject: </b>%s<br>' % cgi.escape(subject_header)
    else:
        title = '(No Subject)'
    base_headers += '<b>Date: </b>%s<br>' % cgi.escape(date_header)

    all_headers = ''
    for item in headers:
        template = '<b>{key}: </b>{value}<br>'
        all_headers += template.format(key=item[0],
                                        value=cgi.escape(item[1]))

    text = ''
    for item in message[0]:
        text += '<p><pre>%s</pre></p>' % cgi.escape(item)

    html = ''
    counter = 0
    for item in message[1]:
        html += '<a href="original-html-%d.html">view original html message</a><br>' % counter

    if attachments:
        counter = 0
        files = '<b>attachments:</b><ul>'
        template = '<li><a href="part-{index}/{link}">{name}</a></li>'

        for item in attachments:
            files += template.format(index=counter,
                                    link=urllib.pathname2url(item[0]),
                                    name=cgi.escape(item[0]))
            counter += 1
        files += '</ul>'
    else:
        files = ''

    return message_template.format(title=title,
                                    base_headers=base_headers,
                                    all_headers=all_headers,
                                    text=text,
                                    html=html,
                                    files=files)


def generate_list_of_messages(msg_list, title=''):
    """Generate a html representation of the message list."""
    table = ''
    template = """
    <tr>
        <td><a href="{id}/message.html">{sender}</a></td>
        <td><a href="{id}/message.html">{subject}</a></td>
        <td nowrap><a href="{id}/message.html">{date}</a></td>
    </tr>
    """

    for item in msg_list:
        table += template.format(id=item[0],
                                sender=cgi.escape(item[1]),
                                subject=cgi.escape(item[2]),
                                date=time.strftime('%Y-%m-%d %H:%M', item[3]))

    return list_template.format(title=cgi.escape(title), table=table)


def get_search_string():
    """Construct search string."""
    search_str = ''
    if options.mindate:
        try:
            mindate = time.strptime(options.mindate, '%Y-%m-%d')
            search_str += ' BEFORE %s' % time.strftime('%d-%b-%Y', mindate)
        except ValueError:
            logging.critical('Unable to parse date, exit')
            sys.exit(1)
    if options.minsize > 0:
        search_str += ' LARGER %d' % (options.minsize * 1024)
    if search_str == '':
        search_str = 'ALL'
    search_str = '(' + search_str.strip() + ')'

    return search_str


def main():
    processed = []
    uids = []

    process_options()

    if options.debug:
        logging.basicConfig(level=logging.DEBUG)

    if not os.path.isdir(options.outputdir):
        try:
            os.mkdir(options.outputdir)
        except OSError as (errno, strerror):
            print 'Unable to create %s directory: %s' % (options.outputdir, strerror)
            sys.exit(1)

    search_str = get_search_string()
    logging.debug('IMAP search string is %s' % search_str)

    try:
        if options.ssl:
            imap = IMAP4_SSL(options.server)
        else:
            imap = IMAP4(options.server)
        imap.login(options.user, options.password)
        imap.select()

        typ, msg_nums = imap.search(None, search_str)

        # fetch messages uids
        for num in msg_nums[0].split():
            typ, msg = imap.fetch(num, '(UID)')
            if typ != 'OK':
                logging.warning('Unable to fetch uid, skipping message %s', num)
                continue
            m = re.match(r".*UID\s+(\d+).*", msg[0])
            uids.append(m.groups()[0])

        for uid in uids:
            typ, data = imap.uid('FETCH', uid, '(RFC822)')
            logging.debug('Fetch message with uid %s' % uid)

            try:
                headers = process_message(uid, data[0][1])
            except:
                logging.warning('Unable to process message with uid %s: %s' % (uid, sys.exc_info()[1]))
            else:
                subject_header = extract_header('Subject', headers)
                if not subject_header:
                    subject_header = '(No Subject)'
                date_struct = email.utils.parsedate(extract_header('Date', headers))
                if not date_struct:
                    logging.warning('Unable to parse date for message with uid %s' % uid)
                    date_struct = time.gmtime(0)
                processed.append((uid,
                                extract_header('From', headers),
                                subject_header,
                                date_struct))

        process_overviews(processed)

        # remove processed messages
        if options.remove:
            for item in processed:
                imap.uid('STORE', item[0], '+FLAGS', '(\\Deleted)')
                logging.debug('Delete message with uid %s' % item[0])
            imap.expunge()

    except socket.error, e:
        logging.critical('Unable to connect to the IMAP server: %s' % str(e))
        sys.exit(1)
    except IMAP4.error, e:
        logging.critical('IMAP4 Error: %s' % str(e))
        sys.exit(1)
    except OSError as (errno, strerror):
        logging.critical('OS error({0}): {1}'.format(errno, strerror))
        sys.exit(1)
    except IOError as (errno, strerror):
        logging.critical('IO error({0}): {1}'.format(errno, strerror))
        sys.exit(1)

    imap.close()
    imap.logout()


if __name__ == '__main__':
    main()
