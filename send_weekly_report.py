from send_email import send_email
from xlrd import xldate_as_tuple
import datetime
import xlrd
import os


# 输入Email地址和口令:
from_addr = 'chenjiefeng@forcegames.cn'
password = 'Trxi900802'
# 输入SMTP服务器地址:
smtp_server = 'smtp.exmail.qq.com'
# 输入收件人地址：
#to_addr = ['chenjiefeng@forcegames.cn']
to_addr = ['chenjiefeng@forcegames.cn', '41816456@qq.com']

# 邮件主题
today = datetime.datetime.now().strftime("%Y%m%d")
email_header = '周报_陈捷丰_' + today
# 邮件正文
readbook = xlrd.open_workbook(r'E:\工作\4.工作安排\2019CMDB需求计划.xlsx')
today2 = datetime.datetime.now().strftime("%m%d")
delta = datetime.timedelta(days=7)
next = (datetime.datetime.now() + delta).strftime("%m%d")
sheet1 = readbook.sheet_by_name(today2)
sheet2 = readbook.sheet_by_name(next)
nrows1 = sheet1.nrows
ncols1 = sheet1.ncols
nrows2 = sheet2.nrows
ncols2 = sheet2.ncols

html_header = '<html><head></head><body><p>Hi，小龙</p><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;以下是我周报内容，请查收：</p><p>本周工作：</p>'
this_week = ''
thead = ''
for i in range(0, ncols1):
	thead = thead + '<th>' + sheet1.cell(0,i).value + '</th>'
thead = '<tr>' + thead + '</tr>'

tbody = ''
for i in range(1, nrows1):
	for j in range(0, ncols1):
		cell = sheet1.cell_value(i,j)
		ctype = sheet1.cell(i,j).ctype
		if ctype == 3:
			date = xldate_as_tuple(cell, 0)
			date = datetime.datetime(*date)
			cell = date.strftime('%Y-%m-%d')
		if ctype == 2 and cell % 1 == 0:
			cell = int(cell)
		tbody = tbody + '<td>' + str(cell) + '</td>'
	tbody = '<tr>' + tbody + '</tr>'
this_week = '<table border="1" cellspacing="0">' + thead + tbody + '</table>'


next_week = ''
thead = ''
for i in range(0, ncols2):
	thead = thead + '<th>' + sheet2.cell(0,i).value + '</th>'
thead = '<tr>' + thead + '</tr>'

tbody = ''
for i in range(1, nrows2):
	for j in range(0, ncols2):
		cell = sheet2.cell_value(i,j)
		ctype = sheet2.cell(i,j).ctype
		if ctype == 3:
			date = xldate_as_tuple(cell, 0)
			date = datetime.datetime(*date)
			cell = date.strftime('%Y-%m-%d')
		if ctype == 2 and cell % 1 == 0:
			cell = int(cell)
		tbody = tbody + '<td>' + str(cell) + '</td>'
	tbody = '<tr>' + tbody + '</tr>'
next_week = '<table border="1" cellspacing="0">' + thead + tbody + '</table>'

html_foot = '<p>其他事项：</p><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;无</p></html>'
html = html_header + this_week + '<p>下周安排：</p>' + next_week + html_foot
email_content = html


	
send_email(from_addr, password, smtp_server, to_addr, email_content, email_header)