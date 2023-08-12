# -*- coding = utf-8 -*-
# @Time : 2021/11/27 14:15
# @Author : 谢扬筱
# @File : demo1.py

import numpy as np
import pandas as pd

#创建序列
s = pd.Series([1, 2, 3, np.nan, 444])
print(s)
#
datas=pd.date_range("20211127",periods=7)
print(datas)
#使用DataFrame,DataFrame，它是一个面向列（column-oriented）的二维表结构，且含有 行标和列标
df=pd.DataFrame(np.random.randn(7,2),index=datas,columns=["a","b"])
print(df)
print(df.dtypes)
df2=pd.DataFrame(np.arange(12).reshape(3,4))
print(df2)
print(df2.dtypes)
