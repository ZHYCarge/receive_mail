# -*- coding: utf-8 -*-
# encoding = utf8
# @Time    : 2022/6/12 17:57
# @Author  : ZHYCarge
# @Email   : ZHYCarge@126.com
# @File    : main.py
# @Brief   : 如下：
# 实现内容：
# 将邮件附件进行下载到本地，并将指定邮件标记为已读

import imaplib
import email
import logging
import os
import configparser

# setting
con = configparser.ConfigParser()
con.read('./config.ini', encoding='utf-8')
items = con.items('mail_box')
mail_box = dict(items)
items = con.items('box_list')
box_list = dict(items)
items = con.items('principal')
principal = dict(items)


## 设置日志情况
if mail_box['log_level'] == 'INFO':
    logging.basicConfig(format='%(asctime)s - line:%(lineno)d] - %(levelname)s: %(message)s',
                    filename='./run.log',
                    filemode='a',
                    level=logging.INFO)

elif mail_box['log_level'] == 'DEBUG':
    logging.basicConfig(format='%(asctime)s - line:%(lineno)d] - %(levelname)s: %(message)s',
                    filename='./run.log',
                    filemode='a',
                    level=logging.DEBUG)
    logging.debug('您已进入测试环境中，请留意')

## 整体设置
mailbox = []
for a in box_list:
    mailbox.append(box_list[a])
principals = []
for a in principal:
    principals.append(principal[a])


# 获得邮件主题（返回邮件主题【str】）
def Get_title(message):
    msgCharset = email.header.decode_header(message.get('Subject'))[0][1]  # 获取邮件标题并进行进行解码，通过返回的元组的第一个元素我们得知消息的编码
    title = email.header.decode_header(message.get('Subject'))[0][0].decode(msgCharset)  # 获取标题并通过标题进行解码
    return title


#判断文件夹是否存在于目录中
def Judge_folder(dir):
    if os.path.exists(dir):
        return
    logging.info(f'【{dir}】目录不存在，已进行创建')
    os.mkdir(dir)


# 获取邮箱的附件
def Get_file(message,locate):
    dir = "./"+locate+"/"
    Judge_folder(dir)
    for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename_unchar = part.get_filename()
        fne = email.header.decode_header(filename_unchar)
        filename = fne[0][0].decode(fne[0][1])
        logging.info(f'获取到附件[{filename}]，正在尝试下载')
        attach_data = part.get_payload(decode=True)
        f = open(dir+filename, 'wb') #注意一定要用wb来打开文件，因为附件一般都是二进制文件
        f.write(attach_data)
        f.close()
        logging.info(f'[{filename}]附件下载成功!')
        return
    logging.info('未获取到附件')



# 标记邮件（需要在select中可以修改邮件）
def Set_flags(uid,conn,status,title):
    if status == '已读':
        s = '+'
    else:
        s = '-'

    typ, _ = conn.store(uid, s+'FLAGS', '\\Seen')
    if typ == 'OK':
        logging.debug(f'已经将邮件【{title}】标记为{status}')



# 邮箱登录
def Login():
    conn = imaplib.IMAP4_SSL(mail_box['mail_ssl'])
    logging.info(f"'已连接服务器{mail_box['mail_ssl']}'")
    conn.login(mail_box['mail_user'], mail_box['mail_password'])
    logging.info(f'已登陆{mail_box["mail_user"]}账户')
    imaplib.Commands['ID'] = ('AUTH')
    args = ("name", "ZHYCarge", "contact", mail_box['mail_user'], "version", "1.0.0", "vendor", "myclient")
    conn._simple_command('ID', '("' + '" "'.join(args) + '")')
    return conn

# 查询邮件文件夹
def BoxList(conn):
    for i in conn.list()[1]:
        logging.debug(i)


# 邮件搜索
def mail_Seach(conn):
    for mail_boxs in mailbox:
        conn.select(mailbox=mail_boxs,readonly=False)    # True表示只读取文件，Flase 表示可以对邮件进行更改
        typ,num = conn.search(None,mail_box['read_mail']) # UNSEEN 表示邮件未读 SEEN表示邮件已读  ALL代表全部
        if typ == 'OK':
            logging.info("成功搜索到邮箱文件夹内信息")
        else:
            logging.error('邮件文件夹信息获取失败！请重试！')
        for uid in num[0].split():
            typ,data = conn.fetch(uid,"(RFC822)")
            if typ == 'OK':
                text = data[0][1].decode("utf-8")
                message = email.message_from_string(text)  # 转换为email.message对象
                title = Get_title(message)
                logging.info(f'获取到主题为【{title}】的邮件')
                if title.split('-')[3] in principals:
                    logging.debug('邮件符合配置材料负责人，进行尝试下载附件')
                    locate = title.split('-')[2]
                    Get_file(message,locate)
                    Set_flags(uid,conn,'已读',title)
                else:
                    logging.debug('邮件不归配置负责人所管，将其重新标记为未读')
                    Set_flags(uid,conn,'未读',title)
            else:
                logging.error(f"邮件编号'{uid}'信息获取失败，请重试！")



# 主函数
if __name__ == '__main__':
    conn = Login()
    BoxList(conn)
    mail_Seach(conn)



