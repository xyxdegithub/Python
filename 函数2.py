# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 16:59
# @Author : 谢扬筱
# @File : 函数2.py

#打印横线
def printHengxian():
    print("-"*30)
#printHengxian()

#打印若干条横线
def printN(n):
    i=1
    while i<=n:
        printHengxian()
        i=i+1

#printN(5)

#求三个数之和
def sum(a,b,c):
    return a+b+c
#sum(1,2,3)

#求三个数的平均数
def pingjun(a,b,c):
    print(sum(a,b,c)//3)
pingjun(1,2,3)