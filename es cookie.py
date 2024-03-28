from http import cookiejar
import requests
from requests.auth import HTTPDigestAuth
import re
import urllib.parse
import os
import urllib
import urllib.request
import requests


# Store the cookies and create an opener that will hold them
cj = cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

# Add our headers
opener.addheaders = [('User-agent', 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.75 Safari/537.36')]

# Install our opener (note that this changes the global opener to the one
# we just made, but you can also just call opener.open() if you want)
urllib.request.install_opener(opener)

# The action/ target from the form
url = 'https://www.esutures.com/account/history/'

# Input parameters we are going to send
payload = {'action': 'signIn',
'login_id': 'mdujowich@vrcvet.com',
'pass': '1820Monterey'}

# Use urllib to encode the payload
data = urllib.parse.urlencode(payload)
data=data.encode("utf-8")
# Build our Request object (supplying 'data' makes it a POST)
req = urllib.request.Request(url, data)

# Make the request and read the response
resp = urllib.request.urlopen(req)

content=resp.read()
print(content)
