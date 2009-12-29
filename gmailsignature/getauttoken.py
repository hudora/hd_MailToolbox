from httplib2 import Http
from urllib import urlencode

h = Http()
data = {"Email": "j.otten@hudora.de", "Passwd": "Hu#14jo", "accountType": "HOSTED", "service": "apps"}

resp, content = h.request("https://www.google.com/accounts/ClientLogin", "POST", urlencode(data), headers={'Content-Type': "application/x-www-form-urlencoded"})

print "This is the authentication token you need:"
print resp, content
