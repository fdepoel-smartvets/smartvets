import requests
from xml.etree import ElementTree

url="https://soap.e-boekhouden.nl/soap.asmx?wsdl"
#headers = {'content-type': 'application/soap+xml'}
headers = {'content-type': 'text/xml'}
body = """<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <OpenSession xmlns="http://www.e-boekhouden.nl/soap">
      <Username>bbrocks</Username>
      <SecurityCode1>c872714821477aab2979b01eddaabde8</SecurityCode1>
      <SecurityCode2>A5700B10-C51A-4F9A-88C0-F4BF01F303EE</SecurityCode2>
    </OpenSession>
  </soap12:Body>
</soap12:Envelope>"""

response = requests.post(url,data=body,headers=headers)
tree = ElementTree.fromstring(response.content)
print(response.content)

# define namespace mappings to use as shorthand below
namespaces = {
    'soap': 'http://www.w3.org/2003/05/soap-envelope',
    'a': 'http://www.e-boekhouden.nl/soap',
}

# reference the namespace mappings here by `<name>:`
names = tree.findall(
    './soap:Body'
    '/a:OpenSessionResponse'
    '/a:OpenSessionResult'
    '/a:SessionID',
    namespaces,
)
for name in names:
    session_id = name.text
print(session_id)

body = """<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <GetFacturen xmlns="http://www.e-boekhouden.nl/soap">
      <SessionID>{0}</SessionID>
      <SecurityCode2>A5700B10-C51A-4F9A-88C0-F4BF01F303EE</SecurityCode2>
      <cFilter>
        <Factuurnummer>F00317</Factuurnummer>
      </cFilter>
    </GetFacturen>
  </soap12:Body>
</soap12:Envelope>"""
body = body.format(session_id)
response = requests.post(url,data=body,headers=headers)
print(response.content)

