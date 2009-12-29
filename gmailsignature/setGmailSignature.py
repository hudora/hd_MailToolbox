#/usr/bin/python
# encoding: utf-8

import httplib2
import sys

class Signature(object):
    def __init__(self, users, domain, apikey):
        self.users = users
        self.domain = domain
        self.apikey = apikey
        self.signature = r"""HUDORA GmbH, Jaegerwald 13, 42897 Remscheid, Germany -  http://www.hudora.de/&#010;
Amtsgericht Wuppertal, HRB 12150, UStId: DE 123241519&#010;
Geschaeftsfuehrer: Evelyn Dornseif, Dr. Maximillian Dornseif, Aufsichtsrat: Eike Dornseif"""

        print [self.signature]

    def setSignature(self, user):
        url = "https://apps-apis.google.com/a/feeds/emailsettings/2.0/%s/%s/signature" % (self.domain, user)

        print url

        body = """<?xml version="1.0" encoding="utf-8"?>
<atom:entry xmlns:atom="http://www.w3.org/2005/Atom" xmlns:apps="http://schemas.google.com/apps/2006">
    <apps:property name="%(name)s" value="%(content)s" />
</atom:entry>""" % {'name': 'signature', 'content': self.signature}

        print body

        status, content = httplib2.Http().request(url, "PUT", body, headers={"Content-type": "application/atom+xml",
                "Content-length": str(len(body)),
                "Authorization": "GoogleLogin auth=%s" % self.apikey})

        if status['status'] != "200":
            print "! Error:", status, content


    def setSignatureAllUsers(self):
        for user in self.users:
            self.setSignature(user.strip())


if __name__ == "__main__":
    userlist = open(sys.argv[1]).readlines()
    sig = Signature(userlist, sys.argv[2], sys.argv[3])
    sig.setSignatureAllUsers()
