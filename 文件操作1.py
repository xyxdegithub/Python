# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 17:17
# @Author : 谢扬筱
# @File : 文件操作1.py

#打开test.txt文件,文件不存在就创建
file=open("test.txt","w")
file.write("Hello test.txt,I am here")

file.close()