# -*- codeing = utf-8 -*-
# @Time : 2021/11/3 20:18
# @Author : 谢扬筱
# @File : 循环1.py
cityName="chengdu"
for i in cityName:
    print(i,end="\t")
print()

a=["aa","bb","cc","dd"]
for i in range(len(a)):
    print(i,a[i],end="\t")
print()

#while示例
i=0
while i<5:
    print("这是第%d次循环"%(i+1))
    print(i)
    i=i+1

#1-100求和
sum=0
for i in range(1,101):
    sum=i+sum
print(sum)