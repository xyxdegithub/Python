# -*- coding = utf-8 -*-
# @Time : 2021/11/9 18:45
# @Author : 谢扬筱
# @File : tesrXlwt.py

# 创建一个excel表，写入数据并保存
import xlwt

# 构造方法创建对象
workbook = xlwt.Workbook(encoding="utf-8")
# 创建表
sheet = workbook.add_sheet("sheet1")
# 写入数据,在表格中写入个九九乘法表要搞清楚横纵坐标
for i in range(0, 9):
    for j in range(0, i + 1):
        result = "%d*%d=%d" % (j + 1, i + 1, (i + 1) * (j + 1))
        sheet.write(i, j, result)
# 保存数据
workbook.save("九九乘法表.xls")
# workbook.save("C:\\Users\\xyx\\Desktop\\九九乘法表.xls")
print("运行结束")
