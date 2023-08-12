# -*- coding = utf-8 -*-
# @Time : 2021/11/12 11:16
# @Author : 谢扬筱
# @File : numpy创建矩阵.py

import numpy广播机制 as np
# 使用numpy创建矩阵,把字符串生成矩阵
m1 = np.mat("1,2;5,6")
print(m1)
print(type(m1))  # <class 'numpy.matrix'>
print(m1.ndim)
print(m1.shape)
m2 = np.mat("2,3;4,5")
print(m2)
m = m1 * m2  # 矩阵的乘法：每一行乘以每一列
print(m)

# 生成全为一的矩阵
a1 = np.ones((3, 4), dtype=np.int16)
print(a1)
print(a1.dtype)
# 生成全为0的矩阵
a2 = np.zeros((3, 4), dtype=np.int32)
print(a2)
print(a2.dtype)
#
a3=np.empty((3,4))
print(a3)
print(a3.dtype)
#
a4=np.linspace(1,10,6)
print(a4)
print(a4.dtype)
a5=a4.reshape(3,2)
print(a5)