# -*- codeing = utf-8 -*-
# @Time : 2021/11/5 0:02
# @Author : 谢扬筱
# @File : 列表5.py
import random
names=["裴思宇","叶云龙","李云龙","楚云飞","刘备","刘邦","唐晟宇"]
rooms=[[],[],[]]
#把名单随机分配进每个room
for name in names:
    room=random.randint(0,2)
    rooms[room].append(name)
#下面嵌套循环逻辑要清楚
i=1
for roomNameNumber in rooms:
    print("第%d个教室的人数为%d:"%(i,len(roomNameNumber)))
    i=i+1
    for roomName in roomNameNumber:
        print(roomName,end=" ")
    print()
    print("-"*25)