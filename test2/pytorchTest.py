# -*- coding = utf-8 -*-
# @Time : 2021/12/7 21:44
# @Author : 谢扬筱
# @File : pytorchTest.py
# cat test_gpu.py
from torch.backends import cudnn
import torch

if __name__ == '__main__':
    print(torch.__version__)
    # 测试 CUDA
    print("Support CUDA ?: ", torch.cuda.is_available())
    x = torch.tensor([10.0])
    x = x.cuda()
    print(x)
    y = torch.randn(2, 3)
    y = y.cuda()
    print(y)
    z = x + y
    print(z)
    # 测试 CUDNN
    print("Support cudnn ?: ", cudnn.is_acceptable(x))
