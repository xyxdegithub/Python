# -*- codeing = utf-8 -*-
# @Time : 2021/11/4 9:31
# @Author : 谢扬筱
# @File : 列表1.py
#python中的list可以存储不同类型的数据
#1.对列表进行增
nameList=["李四","王五","张三"]
print(nameList[0])
for name in nameList:
    print(name,end=" ")
print(len(nameList))
#增加list中的数据,在最后增加
nameList.append("陈独秀")
for name in nameList:
    print(name,end=" ")
print(len(nameList))
#拿到元素下标和内容
for i,name in enumerate(nameList):
    print(i,name)


