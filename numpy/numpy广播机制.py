# -*- coding = utf-8 -*-
# @Time : 2021/12/7 19:43
# @Author : 谢扬筱
# @File : numpy广播机制.py
import numpy as np

A=np.arange(0,40,10).reshape(4,1)
B=np.arange(0,3)
print("A矩阵的形状{},B矩阵的形状{}".format(A.shape,B.shape))
print(A)

C=A+B
print("C矩阵的形状{}".format(C.shape))
print(C)