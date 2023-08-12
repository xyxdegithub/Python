# -*- coding = utf-8 -*-
# @Time : 2021/12/7 19:21
# @Author : 谢扬筱
# @File : 循环计算向量点积.py
import time
import numpy as np

x1 = np.random.rand(10000000)
# print(x1)
x2 = np.random.rand(10000000)
# 返回当前进程的系统和用户CPU时间总和的值（以小数秒为单位）作为浮点数
tic = time.process_time()
dot = 0
for i in range(len(x1)):
    dot += x1[i] * x2[i]
toc = time.process_time()
print("dot=" + str(dot) + "\n"
      + "process_time=" + str(1000 * (toc - tic)) + "ms")

# 使用numpy求点积
tic = time.process_time()
dot = 0
dot = np.dot(x1, x2)
toc = time.process_time()
print("dot=" + str(dot) + "\n" + "process_time=" + str(1000 * (toc - tic)) + "ms")
