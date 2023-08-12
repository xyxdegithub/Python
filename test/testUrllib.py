# -*- coding = utf-8 -*-
# @Time : 2021/11/5 23:24
# @Author : 谢扬筱
# @File : testUrllib.py
import urllib.request

response=urllib.request.urlopen("http://www.baidu.com")
#print(response.read().decode())
print(response.getheaders())
print(response.getheader("Server"))


response2=urllib.request.urlopen("http://www.douban.com")
print(response2.read().decode())

