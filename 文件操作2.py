# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 17:25
# @Author : 谢扬筱
# @File : 文件操作2.py

f=open("test.txt","r")

content=f.read(11)
content2=f.read()

print(content)
print(content2)  #第二次开始不在开头读了


f.close()