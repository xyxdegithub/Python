# -*- codeing = utf-8 -*-
# @Time : 2021/11/3 20:53
# @Author : 谢扬筱
# @File : 循环2.py
# break退出循环，continue退出一次循环，这就是区别
n = 0
while n < 10:
    print("*" * 35)
    n += 1
    if n == 5:
        # break
        continue
    print(n)

# 用python打印九九乘法表
for i in range(1, 10):
    for j in range(1, i + 1):
        print("%d*%d=%d" % (j, i, i * j), end="\t")
    print()
print("*************************************************************")
n = 1
while n < 10:
    m = 1
    while m < n + 1:
        print("%d*%d=%d" % (m, n, n * m), end="\t")
        m = m + 1
    n = n + 1
    print()
