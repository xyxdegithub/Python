# -*- codeing = utf-8 -*-
# @Time : 2021/11/4 10:04
# @Author : 谢扬筱
# @File : 列表3.py

# 对列表删除
movieList = ["长津湖", "我和我的祖国", "我和我的父辈", "李焕英"]
print("*********删除前列表数据*********")
for name in movieList:
    print(name)
print("*********第一次删除后列表数据*********")
del movieList[0]  # 根据下标进行删除
for name in movieList:
    print(name)
print("*********第二次删除后列表数据*********")
movieList.remove("我和我的祖国")  # 根据指定内容进行删除
for name in movieList:
    print(name)
print("*********第二次删除后列表数据*********")
movieList.pop()  # 弹出列表最后一个元素
for name in movieList:
    print(name)

# 改列表
print("*********改变列表数据*********")
movieList[0] = "战狼2"
for name in movieList:
    print(name)
