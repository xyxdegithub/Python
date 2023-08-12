# -*- coding = utf-8 -*-
# @Time : 2021/11/12 12:41
# @Author : 谢扬筱
# @File : numpy中的随机数.py
#随机数可以原来生成数组
import numpy广播机制 as np
#没参数生成一个随机数(0-1)浮点数
print(np.random.rand())

print(np.random.rand(3,2,4))

print(np.random.uniform(2,5,4))
print(np.random.randint(2,5,4))
print(np.random.randint(1,5,4))

print(np.random.randn(2,2))   #标准正态分布数组
print(np.random.normal(2,5,4))  #产生正态分布数组
