# -*- coding = utf-8 -*-
# @Time : 2021/11/9 20:25
# @Author : 谢扬筱
# @File : testSqlite3.py
#轻量的数据库，就是一个文件
import sqlite3

conn=sqlite3.connect("test.db")

print("创建连接数据库成功!")

c=conn.cursor()  #获取游标

#创建表
sql='''
    create table student
    (id int primary key not null,
    name text not null,
    age int not null,
    address char(100) not null) ;
'''

sql2='''
insert into student(id,name,age,address) values (1,"王五",20,"上海市浦东新区")
'''

# sql='''
#         select id,name,age,address from student
# '''
# #执行sql语句
# result=c.execute(sql)   #查询后会返回结果
# for row in result:
#     print("id为:",row[0])
#     print("姓名为:",row[1])
#     print("年龄为:",row[2])
#     print("地址为:",row[3])
c.execute(sql)
c.execute(sql2)
conn.commit()
conn.close()
print("查询数据成功!")