# -*- coding = utf-8 -*-
# @Time : 2021/11/12 10:56
# @Author : 谢扬筱
# @File : numpy中的sum方法.py

import numpy创建矩阵 as np
#用numpy创建一个二维数组
a1=np.array([[0,1,2,3],[4,5,6,7],[8,9,10,11]])

print(a1)
print(np.sum(a1))
print(np.sum(a1,axis=0))
print(np.sum(a1,axis=1))
print("*"*80)
#用numpy创建一个三维数组
a2=np.array([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])
a3=a2.reshape(2,2,4)   #两个二维数组，二维数组再两行两列
print(a2)
print(a3)
print(np.sum(a3))
print(np.sum(a3,0))  #两个二维数组上的对应元素相加
print(np.sum(a3,1))
print(np.sum(a3,2))  #对三位数组的第三个4个量加起来