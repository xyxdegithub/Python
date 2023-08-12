# -*- coding = utf-8 -*-
# @Time : 2021/12/7 19:07
# @Author : 谢扬筱
# @File : 批处理数据.py
import numpy as np

data_train=np.random.randn(10000,2,3)

print("data_train的sahpe是:",data_train.shape)
#打乱shuju
np.random.shuffle(data_train)
batch_size=100

for i in range(0,len(data_train),batch_size):
    x_batch_sum=np.sum(data_train[i:i+batch_size])
    print("第{}批次，该批次数据之和为{}".format(int((i/batch_size))+1,x_batch_sum))
