# -*- coding = utf-8 -*-
# @Time : 2021/11/14 15:20
# @Author : 谢扬筱
# @File : 绘图3.py
import matplotlib.pyplot as plt
#设置字体
plt.rcParams["font.family"]="Microsoft YaHei"
plt.figure(figsize=(4,4),facecolor="white")

plt.subplot(221)
plt.title("子图一")

plt.subplot(222)
plt.title("子图二")

plt.subplot(223)
plt.title("子图三")

plt.subplot(224)
plt.title("子图四",color="b")
#设置全局标题
plt.suptitle("全局标题",color="red",fontsize=30)
plt.tight_layout()
plt.show()