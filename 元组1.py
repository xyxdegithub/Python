# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 10:26
# @Author : 谢扬筱
# @File : 元组1.py

#元组中数据不能更改
a=(1,)
print(a)
print(type(a)) #<class 'tuple'>

b=(1,2,3)
c=("战三","李四","王五")
d=b+c
print(d)     #(1, 2, 3, '战三', '李四', '王五')
print(d[-2:-1])  #左闭右开