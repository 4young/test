import requests
from requests.exceptions import RequestException
target = "http://219.166.7.50"
port = "80"
print("%s:%s"%(target,port))
try:
  res = requests.get(target + ':' + port, timeout = 5)
  print(res.text)
except RequestException:
  print("[-]RequestException:timeout")
    
