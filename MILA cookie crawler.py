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
opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36'), ('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'), ('Accept-Encoding','gzip, deflate, br'),\
    ('Accept-Language','en-US,en;q=0.5' ), ("Connection", "keep-alive"), ("Upgrade-Insecure-Requests",'1')]


# Install our opener (note that this changes the global opener to the one
# we just made, but you can also just call opener.open() if you want)
urllib.request.install_opener(opener)

# The action/ target from the form
url = 'https://www.milainternational.com/sales/order/history/'

# Input parameters we are going to send
payload = {'form_key': 'Q9o5HXRrxjW2d5Qp',
'login[username]': 'mvaughan@vrcvet.com',
'login[password]': 'VrccoInventory1820'}

# Use urllib to encode the payload
data = urllib.parse.urlencode(payload)
data=data.encode("utf-8")

session = requests.Session()
response_login = session.post(url, data)
cookies = session.cookies
response_index = session.get(url)
# Build our Request object (supplying 'data' makes it a POST)
#req = urllib.request.Request(url, data)
print(response_index.text)
# Make the request and read the response
#resp = urllib.request.urlopen(req)

#content=resp.read()
#print(content)
