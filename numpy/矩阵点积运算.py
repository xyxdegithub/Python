# -*- coding = utf-8 -*-
# @Time : 2021/12/3 14:50
# @Author : 谢扬筱
# @File : 矩阵点积运算.py
import numpy广播机制 as np

#就是在线性代数学过的矩阵的乘法运算
x1 = np.array([[1, 2], [3, 4]])
x2 = np.array([[5, 6, 7], [8, 9, 10]])
x3 = np.dot(x1, x2)
print(x3)
