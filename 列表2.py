# -*- codeing = utf-8 -*-
# @Time : 2021/11/4 9:47
# @Author : 谢扬筱
# @File : 列表2.py

list=[1,"伞兵",2,"特种兵",3,"装甲兵"]
lenth=len(list)
print(lenth)
n=0
while n<lenth:
    print(list[n],end=" ")
    n=n+1

list1=[4,"爆破"]
list2=[1,2]
list.extend(list1)
print(len(list))
print(list) #[1, '伞兵', 2, '特种兵', 3, '装甲兵', 4, '爆破']
list.append(list2)
print(len(list))
print(list) #[1, '伞兵', 2, '特种兵', 3, '装甲兵', 4, '爆破', [1, 2]]