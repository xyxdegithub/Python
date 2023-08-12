# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 9:57
# @Author : 谢扬筱
# @File : 列表6.py

products = [["iphone",6888],["MacPro",14800],["小米6",2499],["Coffee",31],["Book",60],["Nike",699]]
print("-"*10+"商品列表"+"-"*10)
i=0
for product in products:
    print(i,"\t",product[0],"\t",product[1])
    i=i+1