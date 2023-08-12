# -*- coding = utf-8 -*-
# @Time : 2021/11/11 18:35
# @Author : 谢扬筱
# @File : 爬取网页文字.py

#爬取网页文字一般没写会用代码爬的，直接复制粘贴就行了
import urllib.request
import re
from bs4 import BeautifulSoup

url="https://www.guancha.cn/politics/2021_11_11_614460.shtml"
response=urllib.request.urlopen(url)
#print(response.read().decode())
text=response.read().decode()
bs=BeautifulSoup(text,"lxml")
text2=bs.find_all("div",class_="content all-txt")
#print(text2)

bs2=BeautifulSoup(str(text2),"lxml")
print(bs2.findAll("p"))
# main=re.compile(r'<p(.*?)>(.*?)</p>',re.S)
# result=re.findall(main, str(text2))
# re.sub('')