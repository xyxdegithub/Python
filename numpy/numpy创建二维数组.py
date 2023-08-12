# -*- coding = utf-8 -*-
# @Time : 2021/11/12 10:04
# @Author : 谢扬筱
# @File : numpy创建二维数组.py
#numpy主要是对数组进行计算的库
#二维数组就是一个矩阵
import numpy创建矩阵 as np
#用numpy创建一个二维数组
a1=np.array([[0,1,2,3],[4,5,6,7],[8,9,10,11]])

print(a1)       #输出数组

print(a1.ndim)  #数组的维数

print(a1.shape)  #数组是三行四列的

print(a1.size)  #有12个元素

print(a1[0,0])   #0
#a1[0,0]等于a1[0][0]
print(a1[0][0])  #0

print(a1.dtype)  #不指定默认是int32类型

