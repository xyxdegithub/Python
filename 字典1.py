# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 10:28
# @Author : 谢扬筱
# @File : 字典1.py

#字典键值对，快速查找，键要唯一
dict={"name":"刘亦菲","age":30,"gender":"女"}
print(dict["name"])  #刘亦菲
#在我们不确定字典中是否存在某个键而又想获取其值时，可以使用get方法，还可以设置默认值
print(dict.get("name")) #刘亦菲
print(dict.get("a"))    #None 使用get方法查找不存在的不报错
print(dict.get("a", "不存在"))   #不存在，可以设定返回什么

for key in dict.keys():
    print(key,end=" ")
print()
print("-"*30)
for value in dict.values():
    print(value, end=" ")
print()
print("-" * 30)
#循环遍历键值对
for key,value in dict.items():
    print("key=%s,value=%s"%(key,value))
#使用clear方法清空字典 ，使用del会把变量删掉
dict.clear()
print(dict)