# -*- codeing = utf-8 -*-
# @Time : 2021/11/3 19:18
# @Author : 谢扬筱
# @File : 判断1.py

if(True):
    print(True)
else:
    print(False)

if(0):
    print(True)
else:
    print(False)

score=int(input("请输入你的考试成绩:"))
if score>=90:
    print("你的成绩属于A+")
elif score<90 and score>=80:
    print("你的成绩属于A")
elif score<80 and score>=70:
    print("你的成绩属于B")
elif score<70 and score>=60:
    print("你的成绩属于C")
else:
    print("你的成绩属于D")