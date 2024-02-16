import requests
import json
import pandas as pd

url = "https://testimex.efes.am/webservice/policy"
data = pd.read_json('C:/Users/AramMartirosyan/OneDrive - EFES ICJSC/Desktop/BYBLOS/News/New_Format.json',  orient='index')[0]
data = dict(data)

payload = json.dumps(data)

headers = {
  'Content-Type': 'application/json',
  'Authorization': "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImNsaWVudCI6eyJpZCI6IjI3NyIsIm5hbWUiOiJJa2h0c3lhbmRyIn0sImVudiI6IlBST0QifSwiaXNzIjoid3d3LmltZXguYW0iLCJpYXQiOjE3MDgwODEwNjYsImV4cCI6MTcwODI1Mzg2Nn0.ZopcOp386o-Eb9zVpU5QtwX8M9hHhlWWlH3xxg1h-DM"
}

response = requests.request("POST", url, headers=headers, data=payload)
if response.status_code == 200:
    print("Success")
    result = json.loads(response.text)
    print(result)
else:
    print("Error")
    print(response.text, response.status_code)



