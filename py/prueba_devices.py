import http.client

conn = http.client.HTTPSConnection("restapi.firmar.online")
payload = ''
headers = {
  'Api-Key': 'de637e97-0dae-4d2b-83db-aa11b7ecf016',
}
conn.request("GET", "/SignFromApp/v40//Devices",  payload, headers)
res = conn.getresponse()
data = res.read()
print(data.decode("utf-8"))