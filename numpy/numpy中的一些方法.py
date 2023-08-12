# -*- coding = utf-8 -*-
# @Time : 2021/11/12 10:30
# @Author : 谢扬筱
# @File : numpy中的一些方法.py

import numpy创建矩阵 as np
#使用arange方法
a1=np.arange(4)
print(a1) #[0 1 2 3]

a2=np.arange(0,2,0.2)
print(a2)  #[0.  0.2 0.4 0.6 0.8 1.  1.2 1.4 1.6 1.8]

#使用ones方法创建元素全为1的数组
a3=np.ones(5,dtype=np.int32)
print(a3)

a4=np.ones((3,3),dtype=np.float64)
print(a4)
print(a4.dtype)
