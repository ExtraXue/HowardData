# -*- coding: utf-8 -*-

import pymysql
conn = pymysql.connect(
    host = '127.0.0.1',
    port = 3306,
    user = 'root',
    passwd = 'ExtraXue123!',
    db = 'email_find',
    charset = 'utf8'
)

def connect_mysql(conn):
    #判断连接是否正常
    conn.ping(True)
    #建立操作游标
    cursor = conn.cursor()
    #设置数据输入输出编码格式
    cursor.execute('set names utf8')
    return cursor

#建立连接游标a
cur = connect_mysql(conn)

#2-添加数据库表头
# ID, 根级目录, 一级目录, 二级目录, 接收时间,
# 发件人, 收件人, 抄送人, 邮件主题, 邮件ID,
# 会话主题, 会话ID, 会话历史记录ID, 邮件内容
cur.execute('''CREATE TABLE IF NOT EXISTS email_box (
        ID INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
        根级目录 VARCHAR(255),
        一级目录 VARCHAR(255),
        二级目录 VARCHAR(255),
	接收时间 VARCHAR(100),
        发件人 VARCHAR(200),
        收件人 VARCHAR(2550),
        抄送人 VARCHAR(2550),
        邮件主题 VARCHAR(255),
        邮件ID VARCHAR(255),
        会话主题 VARCHAR(255),
        会话ID VARCHAR(255),
        会话历史记录ID VARCHAR(2550),
        邮件内容 MEDIUMTEXT
        ) DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci''')

#提交&关闭连接
conn.commit()
conn.close()
print('Connection closed !')