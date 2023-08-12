# -*- coding = utf-8 -*-
# @Time : 2021/11/11 19:57
# @Author : 谢扬筱
# @File : 上下文管理器.py

class File():

    #这种方法命名，表明是python中的特有方法
    def __init__(self,filename,mode):
        print("执行__init__()方法")
        self.filename=filename
        self.mode=mode


    def __enter__(self):
        print("执行__enter__()方法")
        self.f=open(self.filename,self.mode)
        return self.f

    def __exit__(self,*args):
        print("执行__exit__()方法")
        self.f.close()

    with open("t.txt", "r") as f:
        print(f.read())

    print("*"*100)

    with File("t.txt","r") as f:
        print(f.read())

