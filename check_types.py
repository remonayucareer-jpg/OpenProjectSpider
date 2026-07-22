import sys, json, ssl, base64
from urllib.request import Request, urlopen, HTTPSHandler, ProxyHandler, build_opener

api_key = '5c229cdb014a2cfc140dae69462458f2f31d16147f855831aa1890ac0eabd10f'
token = base64.b64encode(f'apikey:{api_key}'.encode('utf-8')).decode('ascii')
headers = {'Authorization': f'Basic {token}'}

req = Request('https://pmo.cemsmart.com/api/v3/users', headers=headers)
ctx = ssl._create_unverified_context()
handlers = [ProxyHandler({}), HTTPSHandler(context=ctx)]
opener = build_opener(*handlers)
resp = opener.open(req, timeout=30)
body = resp.read()
data = json.loads(body.decode('utf-8'))
with open('types_output.txt', 'w', encoding='utf-8') as f:
    for t in data.get('_embedded', {}).get('elements', []):
        f.write(f"ID: {t.get('id')}, Name: {t.get('name')}, Login: {t.get('login')}\n")
print("Done, check types_output.txt")
