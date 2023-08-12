# -*- coding = utf-8 -*-
# @Time : 2021/11/12 12:10
# @Author : 谢扬筱
# @File : 矩阵的装置和逆.py

import numpy广播机制 as np
m1=np.mat("[1,2];[3,3]")
print(m1)
print(m1.T)  #矩阵的转置
print(m1.I)  #矩阵求逆
print(m1*m1.I) #一个矩阵乘以他的逆等于单位阵