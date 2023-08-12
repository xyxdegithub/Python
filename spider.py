# -*- coding = utf-8 -*-
# @Time : 2021/12/9 20:39
# @Author : 谢扬筱
# @File : spider.py
import urllib.request
from bs4 import BeautifulSoup
import xlwt

head = {
    # 用户代理信息告诉网页是用浏览器访问
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
}

def first():
    url = "http://quotes.money.163.com/trade/lsjysj_601398.html"

    request = urllib.request.Request(url, headers=head)
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode()
        # 爬取网页源码
        # f=open("data\\网页源码.html","a",encoding="utf-8")
        # f.write(html)
        #print(html)
        bs = BeautifulSoup(html, "lxml")
        # 查找历史年份
        years_select = bs.find("select", {"name": "year"})
        # print(years_select)
        all_years = years_select.find_all("option")
        # print(all_years)
        return all_years
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)


def second():
    all_years = first()
    txt = ""
    for year in all_years:
        year = year.text
        print("正在爬取"+str(year)+"年数据")
        for season in range(4, 0, -1):
            url = "http://quotes.money.163.com/trade/lsjysj_601398.html?year=" + \
                year + "&season=" + str(season)
            request = urllib.request.Request(url, headers=head)
            response = urllib.request.urlopen(request)
            html = response.read().decode()
            bs = BeautifulSoup(html, "lxml")
            dataTable = bs.find("table",
                                {"class": "table_bg001 border_box limit_sale"})
            # 得到每一行的数据,data是每一个季度的数据
            data = dataTable.find_all("tr")
            # print(data)
            # print(len(data))
            for day in data:
                if(len(data) > 1):
                    day_data = day.find_all("td")
                    # f=open("data\\day_data.txt","a")
                    # f.write(str(day_data))
                    for i in day_data:
                        # 每个格子数据之间空格
                        txt = txt + i.text + "\t"
                    if(len(day_data) > 0):
                        # 每天的数据分行输出
                        txt = txt + "\n"
    f = open("data\\data.txt", "a")
    f.write(txt)
    

def save():
    path = "data\\data.txt"
    f = open(path, encoding="utf-8")
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("历史数据", cell_overwrite_ok=True)
    col = (
        "日期",
        "开盘价",
        "最高价",
        "最低价",
        "收盘价",
        "涨跌额",
        "涨跌幅(%)",
        "成交量(手)",
        "成交金额(万元)",
        "振幅(%)",
        "换手率(%)")
    # 把列名写入
    for i in range(0, 11):
        # 设置表格列宽
        sheet.col(i).width = 256 * 12
        sheet.write(0, i, col[i])

    x = 1
    while True:
        # 按行循环，读取文本文件
        line = f.readline()
        if not line:
            break  # 如果没有内容，则退出循环
        for i in range(len(line.split("\t"))-1):
            # print(line.split("\t"))
            item = line.split("\t")[i]
            if(i==0 or i==7 or i==8):
                sheet.write(x, i, item)
            else:
                item=float(item)
                sheet.write(x,i,item)
        x += 1  # excel另起一行
    f.close()
    workbook.save("data\\bank.xlsx")  # 保存xlsx文件


if __name__ == '__main__':
    first()
    second()
    save()
    print("保存数据成功!")
