import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
import json

#data= pd.read_excel('C:\\Users\\inven\\OneDrive\\Desktop\\Transition Order Planning.xlsx',converters={'SKU':str})

# prod={}
# for index, row in data.iterrows():
#     if pd.notnull(row['SKU']):
#         prod[row['Urgent Care/Westside']]=row['SKU']

# session = requests.Session()
# login_url = 'https://store.mwiah.com/sign-in?returnUrl=%2f'
# credentials = {'username': 'MVaughan@vrcvet.com', 'password': 'VrccoInventory1'}
# session.post(login_url, data=credentials)
# url="https://store.mwiah.com/api/mwi/products/search"
# headers = {
#     'Content-Type': 'application/json;charset=UTF-8',
#     'Accept': 'application/json, text/plain, */*',
#     'Referer': 'https://store.mwiah.com/product/product-search',
#     # 可能还需要其他头部，比如认证信息
# }
# #
# # for key,value in prod.items():
# #     response = session.get("https://store.mwiah.com/product/product-search?term={}&size=100&sortKey=search%3A%3Amostrelevant&categoryId=08d09b31-0739-4231-9a52-a59200bf7f1a".format(value))
#
#
# # print(response.text)
# # soup = BeautifulSoup(response.text, 'html.parser')
# # item=soup.findall('span', title='SKU')
#
# #    print(response.text)
#
#


# test a product, but fail while connect with the server. error code:500. server conncetion failure. need API manule.
session = requests.Session()
login_url = 'https://store.mwiah.com/sign-in?returnUrl=%2f'
credentials = {'username': 'MVaughan@vrcvet.com', 'password': 'VrccoInventory1'}
url="https://store.mwiah.com/api/mwi/products/search"
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
b=session.post(login_url, json=credentials,headers=headers)
print(b)
headers = {
    'Accept':'application/json, text/plain, */*',
'Accept-Encoding':'gzip, deflate, br',
'Accept-Language':'en-US,en;q=0.9',
'Connection':'keep-alive',
'Content-Length':'221',
'Content-Type':'application/json;charset=UTF-8',
'Cookie':'SC_ANALYTICS_GLOBAL_COOKIE=fa7b0edd4ced40f39da71c7c0e2f97d2|False; _fbp=fb.1.1673390084040.1064828205; brand-site=redirect=https%3a%2f%2fstore.mwiah.com%2fredirect&login=https%3a%2f%2fstore.mwiah.com%2fsign-in; _hjSessionUser_2298244=eyJpZCI6IjQ2YWI0ZGY2LWNkMzMtNWVlMC1hNTI4LWU4YjlkNGNkMGY0NCIsImNyZWF0ZWQiOjE2NzQwNjQ4ODI2MTgsImV4aXN0aW5nIjp0cnVlfQ==; _ga_1NQ205XCZT=GS1.2.1697744915.8.0.1697744915.0.0.0; ASP.NET_SessionId=scqpplchwkwanjddjuifqdvb; SID=w3; _ga_F05X7X9K4V=GS1.1.1700165831.2.0.1700165834.57.0.0; _gid=GA1.2.425993519.1700508423; .ASPXAUTH=3D91C4110A5FE94B470D299525C69C6D8F3FF1071B5F9135CB6C7932EFAB05880878C11FF02055AFF4BB598FA3C42863C5E246628F3DFA0EFA32CA6A154E31F88133F0DEEA8220E7478367243DDD0B275C2CFAFD9830E92F2953FB99D2241D9F6FBA2D77EA2844E061BCDD3D4BAEDD2131116A45DD1528997E4F67E73F3196D7C8B468335D56BF3793BC1455021678F1; _ga=GA1.2.1914253386.1673390084; _ga_4YN9N4E76Z=GS1.1.1700516677.441.1.1700519633.0.0.0',
'Host':'store.mwiah.com',
'Newrelic':'eyJ2IjpbMCwxXSwiZCI6eyJ0eSI6IkJyb3dzZXIiLCJhYyI6IjI5ODEyNDMiLCJhcCI6IjYwMTM4NzkyNSIsImlkIjoiOGI5NzhmOWI3MzQ4YWNkNSIsInRyIjoiMzY4ZjEwZmVhYWRjODQyMWU4NWJmNzM0MTBmYTc4MDAiLCJ0aSI6MTcwMDUxOTYzMzczNiwidGsiOiIyODM2NzA0In19',
'Origin':'https://store.mwiah.com',
'Referer':'https://store.mwiah.com/product/product-search?term=282932&size=25&sortKey=search%3A%3Amostrelevant&categoryId=08d09b31-0739-4231-9a52-a59200bf7f1a',
'Sec-Ch-Ua':'"Google Chrome";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
'Sec-Ch-Ua-Mobile':'?0',
'Sec-Ch-Ua-Platform':'"Windows"',
'Sec-Fetch-Dest':'empty',
'Sec-Fetch-Mode':'cors',
'Sec-Fetch-Site':'same-origin',
'Traceparent':'00-368f10feaadc8421e85bf73410fa7800-8b978f9b7348acd5-01',
'Tracestate':'2836704@nr=0-1-2981243-601387925-8b978f9b7348acd5----1700519633736',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
'X-Newrelic-Id':'Vg8PUFRXCxAFUVhSBgECUlA='
}
null=None
data = {"accountId":"232197","searchTerm":"282932","page":1,"pageSize":25,"facets":{},"categoryId":"08d09b31-0739-4231-9a52-a59200bf7f1a","searchCategoryMaxHierarchy":6,"sortValue":"search::mostrelevant","restrictionGroup":null}

a=session.post(url,headers=headers,json=data)

print(a)

# if response.status_code == 200:
#     products = response.json().get('results', [])

#     # 提取每个产品的itemId和itemPrice
#     for product in products:
#         item_id = product.get('itemId', 'N/A')
#         item_price = product.get('calculatedPricing', {}).get('itemPrice', 'N/A')
#         print(f"Item ID: {item_id}, Price: {item_price}")

# elif response.status_code == 500:
#     print("Server Error")
#     print("Response content:", response.text)