# -*- codeing = utf-8 -*-
# @Time : 2021/11/4 23:32
# @Author : 谢扬筱
# @File : 列表4.py

#在列表查找特定内容
findName=input("请输入你的名字:")
nameList=["叶云龙","裴思宇","李云龙","楚云飞"]
if findName in nameList:
    print("成功查找")
else:
    print("查找失败,没有你的名字")

#在指定范围查找
list=["a","b","a","d","e","y"]
#在下标1到4范围查找a，如有则返回找到的下标
print(list.index("a", 1, 4))  #都是左闭右开,random是左闭右闭
#计数"a"出现几次
print(list.count("a"))

#列表翻转，排序方法
b=[1,2,4,31,11,0,-1]
print(b)
b.reverse()
print(b)
#默认升序排列
b.sort()
print(b)
#降序排列
b.sort(reverse=True)
print(b)