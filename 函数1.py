# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 15:46
# @Author : 谢扬筱
# @File : 函数1.py

#定义函数
def printInfo():
    print("-"*35)
    print("-"*8+"人生苦短，我用Python"+"-"*8)
    print("-"*35)

#调用函数函数
printInfo()

#定义返回值函数
def divide(a,b):
    shang=a//b
    yushu=a%b
    return shang,yushu
shang,yushu=divide(5,2)
print("商为:%d 余数为:%d"%(shang,yushu))