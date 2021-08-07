# -*- coding:UTF-8 -*-
# author NWU_dx

import poplib
import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer,Image,Table,TableStyle


def line_feed(content):
    re = ""
    i = 0
    length = 60
    for ch in content:
        if ch == '\r':
            continue
        if ch == '\n':
            i = 0
            re = re + '\n'
        else:
            if '\u4e00' <= ch <= '\u9fff':   #中文
                i = i + 2
            else:
                i = i + 1
            re = re + ch
        if i >= length:
            i = 0
            re = re + '\n'
    return re


def create_pdf(content_dict, group):
    nowTime_date = time.strftime('%Y%m%d', time.localtime())
    pdfname = 'dailyfile/'+nowTime_date + group +'.pdf'

    story = []
    stylesheet = getSampleStyleSheet()
    normalStyle = stylesheet['Normal']

    curr_date = time.strftime("%Y-%m-%d", time.localtime())

    # 标题：段落的用法详见reportlab-userguide.pdf中chapter 6 Paragraph
    rpt_title = '<para autoLeading="off" fontSize=24 align=center><b><font face="msyh">'+group+'日报%s</font></b><br/><br/><br/></para>' % (
        curr_date)
    story.append(Paragraph(rpt_title, normalStyle))

    # 表格数据：用法详见reportlab-userguide.pdf中chapter 7 Table
    component_data = [['姓名', '日报']]

    for name in content_dict.keys():
        lists = []
        lists.append(name)
        lists.append(line_feed(content_dict[name]))
        component_data.append(lists)



    # 创建表格对象，并设定各列宽度
    component_table = Table(component_data, colWidths=[50, 400])
    # 添加表格样式
    component_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'msyh'),  # 字体
        ('FONTSIZE', (0, 0), (1, 10), 12),  # 字体大小
        # ('SPAN', (0, 0), (3, 0)),  # 合并第一行前三列
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightskyblue),  # 设置第一行背景颜色
        # ('SPAN', (-1, 0), (-2, 0)),  # 合并第一行后两列
        ('ALIGN', (-1, 0), (-2, 0), 'RIGHT'),  # 对齐
        # ('VALIGN', (-1, 0), (-2, 0), 'MIDDLE'),  # 对齐
        ('VALIGN', (0, 0), (2, 8), 'MIDDLE'),  # 所有表格上下居中对齐
        ('LINEBEFORE', (0, 0), (0, -1), 0.1, colors.black),  # 设置表格左边线颜色为灰色，线宽为0.1
        ('TEXTCOLOR', (0, 1), (-2, -1), colors.black),  # 设置表格内文字颜色
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  # 设置表格框线为红色，线宽为0.5
    ]))
    story.append(component_table)

    doc = SimpleDocTemplate(pdfname)
    doc.build(story)
    return pdfname


# 10.26 测试可用
def get_email(email, password, host):
    # connect to pop3 server
    server = poplib.POP3(host)

    # 身份验证
    server.user(email)
    server.pass_(password)

    # 返回邮件总数目和占用服务器的空间大小（字节数）， 通过stat()方法即可
    # print("Mail counts: {0}, Storage Size: {0}".format(server.stat()))

    # 使用list()返回所有邮件的编号，默认为字节类型的串
    resp, mails, octets = server.list()
    # print("响应信息： ", resp)
    # print("所有邮件简要信息： ", mails)
    # print("list方法返回数据大小（字节）： ", octets)

    index = 1
    msg_contents = []
    while index >= 1 and index <= len(mails):
        resp, lines, octets = server.retr(index)
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        msg = Parser().parsestr(msg_content)
        msg_contents.append(msg)
        index += 1
    server.close()
    return msg_contents


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


# 10.26 测试可用 获取邮件发送者昵称
def get_From(msg):
    value = msg.get("From")
    hdr, addr = parseaddr(value)
    name = decode_str(hdr)
    return name


# 10.26 测试可用 解析邮件内容
def get_content(msg):
    if msg.is_multipart():
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            # print(n)
            re = get_content(part)
            if n == 0:
                break
    else:
        content_type = msg.get_content_type()
        if content_type == 'text/plain' or content_type == 'text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            return content
        else:
            return
    return re


# 10.26测试完毕
def send_alerts_mail(to_list):
    mail_host = 'smtp.163.com'
    port = 465
    send_by = 'nsas_daily@163.com'
    password = 'VXDUKEHFKHMCKHRT'
    content = "这是一封提醒邮件，汇总邮件将于22:15发出，请大家及时提交日报"

    m = MIMEText(content, 'plain', 'utf-8')
    m['From'] = send_by
    m['Subject'] = "日报提醒"
    to = to_list[0]
    for i in range(1, len(to_list)):
        print(i)
        to = to + ',' + to_list[i]
    m['to'] = to
    print(to)
    print(to_list)
    try:
        smpt = smtplib.SMTP_SSL(mail_host, port, 'utf-8')
        smpt.login(send_by, password)
        smpt.sendmail(send_by, to_list, m.as_string())
        print('send alert success!')
    except smtplib.SMTPException as e:
        print('send alert Error:', e)


# 参数为一个字典，
def send_summary_mail(member_list, dict_content, group):
    print("send mail")
    doc_name = create_pdf(dict_content, group)

    mail_host = 'smtp.163.com'
    port = 465
    send_by = 'nsas_daily@163.com'
    password = 'VXDUKEHFKHMCKHRT'
    content = "各位老师同学：你们好！附件中是今天本小组的日志汇总，请各位查收。"

    # 创建一个带附件的实例
    message = MIMEMultipart()
    message['From'] = send_by  # 发送者
    # 邮件标题
    subject = '汇总邮件'
    message['Subject'] = Header(subject, 'utf-8')

    # 邮件正文内容
    message.attach(MIMEText(content, 'plain', 'utf-8'))

    # 添加附件2
    part = MIMEApplication(open(doc_name, 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename=doc_name)
    message.attach(part)
    if len(member_list) == 0:
        return
    to = member_list[0]
    for i in range(1, len(member_list)):
        print(i)
        to = to + ',' + member_list[i]
    message['to'] = to

    try:
        smpt = smtplib.SMTP_SSL(mail_host, port, 'utf-8')
        smpt.login(send_by, password)
        smpt.sendmail(send_by, member_list, message.as_string())
        print('success!')
    except smtplib.SMTPException as e:
        print('Error:', e)

    '''
    for to in member_list:
        message['To'] = to
        try:
            smpt = smtplib.SMTP_SSL(mail_host, port, 'utf-8')
            smpt.login(send_by, password)
            smpt.sendmail(send_by, to, message.as_string())
            print('success!')
        except smtplib.SMTPException as e:
            print('Error:', e)
    '''


def dele_mail(mail, password, host):
    server = poplib.POP3(host)

    # 身份验证
    server.user(mail)
    server.pass_(password)

    mail_max = server.stat()[0]
    for i in range(mail_max):
        server.dele(i + 1)
    print(server.stat())
    server.quit()



