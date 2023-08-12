# -*- coding = utf-8 -*-
# @Time : 2021/11/8 20:04
# @Author : 谢扬筱
# @File : testRe.py

#在正则表达式中比较内容前面加r避免被转义
import re

p=re.compile("AA")   #p是正则表达式，可用来去验证其它的字符串
result=p.search("Abc")  #search方法是验证被校验的内容，如注册账号是输入的内容
print(result)        #输出验证之后的结果

result2=p.search("AAA")
result3=re.search("AA","AAA")  #前面是规则，后面是校验的内容
print(result2)
print(result3)

#search和findall的区别
result4=re.search("[A-Z]","ABC")
print(result4)   #<re.Match object; span=(0, 1), match='A'>
result5=re.findall("[A-Z]","ABC")
print(result5)   #['A', 'B', 'C']
result6=re.findall("[A-Z]+","ABCdsEF")
print(result6)   #['ABC', 'EF']

#sub替换
s=re.sub("a","A","aabc")  #找到a用A替换
print(s)   #AAbc
