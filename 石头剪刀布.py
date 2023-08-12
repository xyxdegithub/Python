# -*- codeing = utf-8 -*-
# @Time : 2021/11/3 20:00
# @Author : 谢扬筱
# @File : 石头剪刀布.py
import random

# 石头剪刀布
for i in range(0,5):
    number = int(input("请输入石头(0),剪刀(1),布(2):"))
    if number not in(0,1,2):
        print("输入数字不符合要求")
    else:
        randomNumber = random.randint(0,2)
        print("随机数是%d:"%randomNumber)
        if randomNumber == number:
            print("平手")
        elif randomNumber == 0 and number == 1:
            print("随机数石头(0)胜剪刀(1)")
        elif randomNumber == 0 and number == 2:
            print("随机数石头(0)输布(2)")
        elif randomNumber == 1 and number == 0:
            print("随机数剪刀(1)输石头(0)")
        elif randomNumber == 1 and number == 2:
            print("随机数剪刀(1)胜布(2)")
        elif randomNumber == 2 and number == 0:
            print("随机数布(2)胜石头(0)")
        else:
            print("随机数剪刀(1)输石头(0)")
    print("******************************************")
# elif randomNumber==2 and number==1:
#    print("随机数布(2)输剪刀(1)")