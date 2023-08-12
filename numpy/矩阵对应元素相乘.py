# -*- coding = utf-8 -*-
# @Time : 2021/12/3 14:46
# @Author : 谢扬筱
# @File : 矩阵对应元素相乘.py
import numpy广播机制 as np

#矩阵对应元素相乘
a=np.array([[1,2],[-1,4]])
b=np.array([[2,0],[3,4]])
print(a*b)
print(np.multiply(a,b))
print(a*2)
print(a/2)