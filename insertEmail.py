# -*- coding: utf-8 -*-
import pymysql
# pywin32.com.win32com.client.gencache import EnsureDispatch as Dispatch
from win32com.client.gencache import EnsureDispatch as Dispatch

# 连接MySQL
conn = pymysql.connect(
      host='127.0.0.1',
      port=3306,
      user='root',
      passwd='root',
      db='email_find',
      charset='utf8mb4')
def connect_mysql(conn):
    #判断链接是否正常
    conn.ping(True)
    #建立操作游标
    cursor=conn.cursor()
    #设置数据输入输出编码格式
    cursor.execute("SET NAMES utf8mb4")
    return cursor

#建立链接游标
cur=connect_mysql(conn)

# SQL语句
insert_sql = 'insert into email_box ' \
             '(根级目录,一级目录,二级目录,接收时间,发件人,收件人,抄送人,邮件主题,邮件ID,会话主题,会话ID,会话历史记录ID,邮件内容)' \
             ' values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'

# 连接Outlook
outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
Accounts = mapi.Folders  # 根级目录（邮箱名称，包括Outlook读取的存档名称）
for Account_Name in Accounts:
    print(' >> 正在查询的帐户名称：',Account_Name.Name,'\n')
    Level_1_Names = Account_Name.Folders  # 一级目录集合（与inbox同级）
    for Level_1_Name in Level_1_Names:
        # 首先，向MySQL提交一级目录的邮件
        print(' - 正在查询一级目录：' , Level_1_Name.Name)
        Mail_1_Messages = Level_1_Name.Items  # 一级文件夹的mail合集
        for xx in Mail_1_Messages:  # xx = 'mail'  # 开始查看单个邮件的信息
            Root_Directory_Name_1 = Account_Name.Name # 记录根目录名称
            Level_1_FolderName_1 = Level_1_Name.Name # 记录一级目录名称
            Level_2_FolderName_1 = ''  # 一级目录肯定没有二级目录，顾留为空
            if (hasattr(xx, 'ReceivedTime')):
                ReceivedTime_1 = str(xx.ReceivedTime)[:-6]  # 接收时间
            else:
                ReceivedTime_1 = ''
            if (hasattr(xx, 'SenderName')):  # 发件人
                SenderName_1 = xx.SenderName
            else:
                SenderName_1 = ''
            if (hasattr(xx, 'To')):  # 收件人
                to_to_1 = xx.To
            else:
                to_to_1 = ''
            if (hasattr(xx, 'CC')):  # 抄送人
                cc_cc_1 = xx.CC
            else:
                cc_cc_1 = ''
            if (hasattr(xx, 'Subject')):  # 主题
                Subject_1 = xx.Subject
            else:
                Subject_1 = ''
            if (hasattr(xx, 'EntryID')):  # 邮件MessageID
                MessageID_1 = xx.EntryID
            else:
                MessageID_1 = ''
            if (hasattr(xx, 'ConversationTopic')):  # 会话主题
                ConversationTopic_1 = xx.ConversationTopic
            else:
                ConversationTopic_1 = ''
            if (hasattr(xx, 'ConversationID')):  # 会话ID
                ConversationID_1 = xx.ConversationID
            else:
                ConversationID_1 = ''
            if (hasattr(xx, 'ConversationIndex')):  # 会话记录相对位置
                ConversationIndex_1 = xx.ConversationIndex
            else:
                ConversationIndex_1 = ''
            if (hasattr(xx, 'Body')):  # 邮件内容
                EmailBody_1 = xx.Body[:25536]
            else:
                EmailBody_1 = ''

            # 写入MySQL
            cur.execute(insert_sql,
                        (Root_Directory_Name_1, Level_1_FolderName_1, Level_2_FolderName_1, ReceivedTime_1, SenderName_1,
                         to_to_1, cc_cc_1, Subject_1, MessageID_1, ConversationTopic_1, ConversationID_1, ConversationIndex_1,
                         EmailBody_1))
        # 然后，判断当前查询的一级邮件目录是否有二级目录（若有多级目录，可以参考此段代码)
        if Level_1_Name.Folders: 
            Level_2_Names = Level_1_Name.Folders  # 二级目录的集合（比如，自建目录的子集）
            for Level_2_Name in Level_2_Names:
                print(' - - 正在查询二级目录：' , Level_1_Name.Name , '//' , Level_2_Name.Name)
                Mail_2_Messages = Level_2_Name.Items  # 二级目录的邮件集合
                for yy in Mail_2_Messages:  # xx = 'mail'  # 开始查看单个邮件的信息
                    Root_Directory_Name_2 = Account_Name.Name # 记录根目录名称
                    Level_1_FolderName_2 = Level_1_Name.Name # 记录一级目录名称
                    Level_2_FolderName_2 = Level_2_Name.Name # 记录二级目录名称
                    if (hasattr(yy, 'ReceivedTime')):
                        ReceivedTime_2 = str(yy.ReceivedTime)[:-6]  # 接收时间
                    else:
                        ReceivedTime_2 = ''
                    if (hasattr(yy, 'SenderName')):  # 发件人
                        SenderName_2 = yy.SenderName
                    else:
                        SenderName_2 = ''
                    if (hasattr(yy, 'To')):  # 收件人
                        to_to_2 = yy.To
                    else:
                        to_to_2 = ''
                    if (hasattr(yy, 'CC')):  # 抄送人
                        cc_cc_2 = yy.CC
                    else:
                        cc_cc_2 = ''
                    if (hasattr(yy, 'Subject')):  # 主题
                        Subject_2 = yy.Subject
                    else:
                        Subject_2 = ''
                    if (hasattr(yy, 'EntryID')):  # 邮件MessageID
                        MessageID_2 = yy.EntryID
                    else:
                        MessageID_2 = ''
                    if (hasattr(yy, 'ConversationTopic')):  # 会话主题
                        ConversationTopic_2 = yy.ConversationTopic
                    else:
                        ConversationTopic_2 = ''
                    if (hasattr(yy, 'ConversationID')):  # 会话ID
                        ConversationID_2 = yy.ConversationID
                    else:
                        ConversationID_2 = ''
                    if (hasattr(yy, 'ConversationIndex')):  # 会话记录相对位置
                        ConversationIndex_2 = yy.ConversationIndex
                    else:
                        ConversationIndex_2 = ''
                    if (hasattr(yy, 'Body')):  # 邮件正文内容
                        EmailBody_2 = yy.Body
                    else:
                        EmailBody_2 = ''

                    # 写入MySQL
                    cur.execute(insert_sql, (Root_Directory_Name_2, Level_1_FolderName_2,
                                             Level_2_FolderName_2, ReceivedTime_2,
                                             SenderName_2,to_to_2, cc_cc_2, Subject_2, MessageID_2,
                                             ConversationTopic_2, ConversationID_2, ConversationIndex_2, EmailBody_2))
        else:
            pass

# 结尾
conn.commit()
conn.close()
print ('\n',' >> Done!')