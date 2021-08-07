# -*- coding:UTF-8 -*-
# author NWU_dx

import importlib,sys
import time
import utils
import csv
importlib.reload(sys)
('utf-8')


nickname = []   #

nickname_realname = {}

nickname_email = {}

nickname_group = {}



mail = {"师大移动小组": {}, "NSAS移动小组": {}, "第三组": {}}

summary_dict = {}
summary_nick_dict = {}

with open('configs.csv', 'r', encoding='GBK') as f:                         # name	 mail	 nick_name	 group_id	group_name
    reader = csv.reader(f)
    i = 0
    for row in reader:
        print(row)
        if len(row) == 0:
            break
        if i == 0:
            i += 1
            continue
        nickname.append(row[2])
        nickname_email[row[2]] = row[1]
        nickname_realname[row[2]] = row[0]
        nickname_group[row[2]] = row[4]

        mail[row[4]][row[2]] = "该同学没有提交日报"               # 第几组的某某某没有提交日报
        summary_dict[row[4]].append(row[1])
        summary_nick_dict[row[4]].append(row[2])

Alerts_sign = 0
summary_sign = 0
dele_sign = 1

while 1:
    nowTime_date = time.strftime('%Y%m%d', time.localtime())                # 当前日期 格式为 20201026
    time_now = time.strftime("%H:%M:%S", time.localtime())
    print(time_now)
    #if 1:
    if '21:45:00' <= time_now <= '21:45:40' and Alerts_sign == 0:           # 发送提醒邮件，30分钟后发送汇总邮件
        # 邮件提醒
        print(time_now, end="     ")
        print("开始发送提醒邮件")
        Alerts_sign = 1
        HOST = "pop3.163.com"
        msg_contents = utils.get_email("email", "password", HOST)
        alerts_list = []
        receive = []
        nick = []
        for msg in msg_contents:
            name = utils.get_From(msg)
            receive.append(name)
        for name in nickname:
            if name in receive:
                print(name + "已经发送日报")
            else:
                print("name is "+name)
                nick.append(name)
                alerts_list.append(nickname_email[name])
        utils.send_alerts_mail(alerts_list)
        print("alert list  ", end="   ")
        print(alerts_list)
        print("nick  ", end="   ")
        print(nick)
        # time.sleep(1700)

    time_now = time.strftime("%H:%M:%S", time.localtime())
    # if 1:
    if '22:15:00' <= time_now <= '22:15:40' and summary_sign == 0:              # 发送汇总邮件
        summary_sign = 1
        dele_sign = 0
        HOST = "pop3.163.com"
        msg_contents = utils.get_email("email", "password", HOST)
        for msg in msg_contents:
            name = utils.get_From(msg)
            content = utils.get_content(msg)
            if name in nickname:
                pass
            else:
                continue

            if mail[nickname_group[name]][name] == "该同学没有提交日报":
                mail[nickname_group[name]][name] = content
        # 将mails中的信息汇总起来
        re = {}
        for group in mail.keys():
            for name in mail[group].keys():
                re[nickname_realname[name]] = mail[group][name]
            print("re list is ", end="  ")
            print(re)
            print("nick list is  ", end="  ")
            print(summary_nick_dict[group])
            utils.send_summary_mail(summary_dict[group], re, group)
            print(re)
            re = {}

        time.sleep(10)
    # break
    # if 1:
    if '22:18:00' <= time_now <= '22:18:40' and dele_sign == 0:  # 初始化
        # 将收件箱邮件移动到其他文件夹
        dele_sign = 1
        HOST = "pop3.163.com"
        utils.dele_mail("email", "password", HOST)

        # 初始化标志变量
        summary_sign = 0
        Alerts_sign = 0

        for group in mail.keys():
            for name in mail[group].keys():
                mail[group][name] = "该同学没有提交日报"

        for item in mail.keys():
            print(mail[item], end="   ")
        break
        time.sleep(84400)
    time.sleep(10)
