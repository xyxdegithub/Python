# -*- codeing = utf-8 -*-
# @Time : 2021/11/3 10:59
# @Author : 谢扬筱
# @File : 输出.py
import keyword

#标准输出
for i in range(0,5):
    print("hello huang")
#查看关键字
print(keyword.kwlist)

#格式化输出
age=22
print("我今年%d岁了"%age)
print("我的名字叫%s，我的国籍是%s"%("小谢","中国"))

print("aaa","bbb","cccc")
print("www","baidu","com",sep=".")
print("Hello",end=" ")
print("World",end="\t")
print("begin",end="\n")
print("end")

#换行输出
print("11111111\n22222222")
