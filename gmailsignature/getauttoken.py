from httplib2 import Http
from urllib import urlencode

h = Http()
data = {"Email": "v.nachname@example.com", "Passwd": "sekrit", "accountType": "HOSTED", "service": "apps"}

resp, content = h.request("https://www.google.com/accounts/ClientLogin", "POST", urlencode(data), headers={'Content-Type': "application/x-www-form-urlencoded"})

print "This is the authentication token you need:"
print resp, content
