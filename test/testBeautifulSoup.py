# -*- coding = utf-8 -*-
# @Time : 2021/11/7 17:44
# @Author : 谢扬筱
# @File : testBeautifulSoup.py

from bs4 import BeautifulSoup
import re

#文档的搜索
file=open("../html.html", "r", encoding="utf-8")  #此文件在项目主目录下，.代表一层目录
html=file.read()
#print(file.read())
#构造函数
bs=BeautifulSoup(html,"html.parser")
print(bs.findAll("a"))
print(bs.find("title").string)
print(bs.title.name)
print(bs.div)
print(bs.div.name)
for item in bs.findAll("a"):
    print(item)
    print(item.get("href"))
    print(item.get("name"))

#用着表达式搜索
print(bs.findAll(re.compile("a")))