import requests
from requests.exceptions import RequestException
target = "http://127.0.0.1"
port = "80"
print("%s:%s"%(target,port))
try:
  res = requests.get(target + ':' + port, timeout = 5)
  print(res.text)
except RequestException:
  print("[-]RequestException:timeout")
    
