import requests

url = 'https://testimex.efes.am/webservice/auth'
user = 'Ikhtsyandr'
password = 'Bb7410/8520*963.-+'

# Provide authentication credentials in the request body
data = {
    'username': user,
    'password': password
}

# Make a POST request to the authentication endpoint
response = requests.post(url, data=data)

# Check the response status
if response.status_code == 200:
    print("Authentication successful")
    # Optionally, you can also print the response content
    print(response.text)
else:
    print("Authentication failed with status code:", response.status_code)
