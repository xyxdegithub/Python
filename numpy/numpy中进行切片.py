# -*- coding = utf-8 -*-
# @Time : 2021/11/12 10:48
# @Author : 谢扬筱
# @File : numpy中进行切片.py

import numpy创建矩阵 as np
#二维数组切片
a=np.array([1,2,3,4,5,6,7,8,9,10,11,12])
a1=a.reshape(3,4)
print(a)
print(a1)
print(a1[1:2])
print(a1[1:2,1:2])
print(a1[:,1:2])