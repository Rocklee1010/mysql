#!/usr/bin/env python3
# _*_ coding: UTF-8 _*_
"""
    读取excel中的每个表的数据 导入到 指定的数据库对应的表中
"""


import pymysql
import xlrd

# 连接数据库excels
db = pymysql.connect(
    host='localhost',
    port=3306,
    user='root',
    password='123456',
    database='excels',
    charset='utf8'
)

# 获取游标
cur = db.cursor()

# 读取excel表格中的表单
'''
1.输入文件名称
2.打开指定路径的文件
3.读取每个表的名称,作为sql的表名
读取每个表里面的第一行的内容,作为表的字段名
'''

# 读取当前文件夹下的excel文件
data = xlrd.open_workbook('./infomations.xlsx')
print('获取对象成功')

# 获取所有的表对象
tables = data.sheets()

# 遍历每个表对象
for sheet in tables:
    # 获取表名称
    table_name = sheet.name

    # 如果此表的行数不为0,则创建数据库表,否则就不创建
    nrows = sheet.nrows

    # 字段长度即列数
    ncols = sheet.ncols

    if nrows != 0:
        # 获取表的字段名称列表
        col_names = [i+' varchar(128)' for i in sheet.row_values(0, start_colx=0, end_colx=None)]

        # sql 语句
        sql = 'create table if not exists %s (%s)charset=utf8' % (table_name, ','.join(col_names))

        # 创建表
        cur.execute(sql)
        db.commit()

        print('创建表成功')

        col_num = ','.join(['%s']*ncols)

        # 从第二行开始进行数据的插入
        sql = 'insert into '+table_name+' values ('+col_num+')'

        # 定一个总列表,里面存储每行数据组成的元组作为一个元素
        data_list = []

        # 读取excel从第二行开始的每行数据,形成列表
        for i in range(1,nrows):
           data_list.append(tuple(sheet.row_values(i, start_colx=0, end_colx=None)))

        try:
            cur.executemany(sql,data_list)
            db.commit()  # 同步数据库
            print('成功插入数据')
        except Exception as e:
            print(e)
            print('插入数据失败')
            db.rollback()

cur.close()
db.close()
