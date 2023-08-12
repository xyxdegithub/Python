# -*- coding = utf-8 -*-
# @Time : 2021/11/5 23:10
# @Author : 谢扬筱
# @File : spider.py

import urllib.request
from bs4 import BeautifulSoup
import re
import xlwt
import sqlite3


def main():
    url = "https://movie.douban.com/top250?start="
    # 爬取网页
    dataList = getData(url)  # 调用getData函数
    # 保存数据到exel表
    # savePath="豆瓣电影Top250.xls"
    # saveData(dataList,savePath)
    # 保存数据到数据库
    dbPath = "test.db"
    saveDataToDB(dataList, dbPath)
# 爬取网页


def getData(url):
    dataList = []
    for i in range(0, 10):
        url2 = url + str(i * 25)  # 进行换页操作，变换网址，找出规律
        html = dUrl(url2)  # 调用dUrl函数

        bs = BeautifulSoup(html, "lxml")
        #print(bs.find_all("div", class_="item"))
        # 用正则表达式进行解析网页
        urlLink = re.compile(r'<a href="(.*?)">')  # 爬取影片详情链接的规则
        pngLink = re.compile(r'<img.*src="(.*?)"', re.S)  # 让包含换行符匹配，这里不加也行
        movieTiltle = re.compile(r'<span class="title">(.*?)</span>')  # 匹配电影标题
        movieRate = re.compile(
            r'<span class="rating_num" property="v:average">(.*?)</span>')
        movieJudgeN = re.compile(r'<span>(\d*)人评价</span>')
        movieIntroduce = re.compile(r'<span class="inq">(.*?)</span>')
        movieDetail = re.compile(r'<p class="">(.*?)</p>', re.S)
        for item in bs.find_all("div", class_="item"):
            # print(item)
            # print("*"*180)
            item = str(item)  # 把item的类型要转换
            data = []

            uLink = re.findall(urlLink, item)[0]
            data.append(uLink)

            pLink = re.findall(pngLink, item)[0]
            data.append(pLink)

            mTiltle = re.findall(movieTiltle, item)
            if len(mTiltle) == 2:
                fTitle = mTiltle[0]
                data.append(fTitle)
                sTitle = mTiltle[1].replace("/", "")
                data.append(sTitle)
            else:
                data.append(mTiltle[0])
                data.append(" ")

            mRate = re.findall(movieRate, item)[0]
            data.append(mRate)

            mJudgeN = re.findall(movieJudgeN, item)[0]
            data.append(mJudgeN)

            mIntroduce = re.findall(movieIntroduce, item)
            if len(mIntroduce) != 0:
                mIntroduce = mIntroduce[0].replace("。", "")
                data.append(mIntroduce)
            else:
                data.append(" ")

            mDetail = re.findall(movieDetail, item)[0]
            mDetail = re.sub('<br(\\s+)?/>(\\s+)?', " ", mDetail)  # 去掉br
            mDetail = re.sub('/', " ", mDetail)  # 替换、
            data.append(mDetail.strip())  # 去掉前后的空格

            dataList.append(data)

    return dataList


# 得到指定网页内容
def dUrl(dUrl):
    head = {
        # 用户代理信息告诉网页是用浏览器访问
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 95.0.4638.69Safari / 537.36"
    }
    request = urllib.request.Request(dUrl, headers=head)
    try:
        reposen = urllib.request.urlopen(request)
        html = reposen.read().decode()
        # print(html)
        return html
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)

# 保存数据,保存到excel表


def saveData(dataList, savePath):
    print("保存数据")
    workbook = xlwt.Workbook(encoding="utf-8")
    sheet = workbook.add_sheet("豆瓣电影Top250")
    col = (
        "电影链接",
        "电影图片链接",
        "电影中文名",
        "电影外国名",
        "电影评分",
        "电影评价人数",
        "电影介绍",
        "电影详情")
    # 把列名写入
    for i in range(0, 8):
        sheet.write(0, i, col[i])
    for j in range(0, 250):
        print("第%d条" % (j + 1))
        data = dataList[j]
        for n in range(0, 8):
            # 把内容写入对应的格子中
            sheet.write(j + 1, n, data[n])
    workbook.save(savePath)
    print("保存数据完毕")

# 初始化数据库


def init_db(dbPath):
    sql = '''
            create table topmovie250
            (
            id integer primary key autoincrement,
            movieLink text,
            pngLink text,
            cname text,
            fname text,
            score real ,
            jnumber real ,
            introduce text,
            detail text
            )

        '''  # 创建数据表
    conn = sqlite3.connect(dbPath)
    c = conn.cursor()
    c.execute(sql)
    conn.commit()
    conn.close()

# 保存数据，保存在sqlite3数据库中


def saveDataToDB(dataList, dbPath):
    init_db(dbPath)
    conn = sqlite3.connect(dbPath)
    c = conn.cursor()

    for data in dataList:
        for index in range(len(data)):
            if index == 4 or index == 5:
                continue
            data[index] = '"' + data[index] + '"'
        # 循环执行250次
        sql = '''
        insert into topmovie250(movieLink,pngLink,cname,fname,score,jnumber,introduce,detail)
                values (%s) ''' % ",".join(data)
        print(sql)
        c.execute(sql)
        conn.commit()
    c.close()
    conn.close()


if __name__ == '__main__':
    main()
    print("爬取结束!")
