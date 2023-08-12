# -*- coding = utf-8 -*-
# @Time : 2021/12/3 14:24
# @Author : 谢扬筱
# @File : test.py
import numpy广播机制 as np

n1=np.linspace(0,2,5)
print(n1)
print(n1.dtype)
print(type(n1))
n2=np.linspace(0,2,5,retstep=True)
print(n2)
print(type(n2))