import requests
from bs4 import BeautifulSoup

def requestsSite (url, param, type):
  res = None
  match type:
    case "get":
      res = requests.get(url, param)
    case "post":
      res = requests.post(url, param)
    case "put":
      res = requests.put(url, param)
    case _:
      print("not allow request type!")
  if res.status_code == 200:
    content = BeautifulSoup(res.content, "html.parser")
    return content
  else:
    print("Request failed -> ", res.status_code)
    return -1