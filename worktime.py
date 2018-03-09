# !/usr/bin/python
# coding:utf-8
# xlsxwriter的基本用法
import xlsxwriter
import datetime, calendar
import sys

year = 2018
month = 1
num_days = calendar.monthrange(year, month)[1]
days = [str(month) + '/' + str(day) for day in range(1, num_days + 1)]

# 1. 建立一個Excel文件
workbook = xlsxwriter.Workbook('demo1.xlsx')

# 2. 建立一個工作表sheet物件
worksheet = workbook.add_worksheet()

# 4. 定義一個加粗的格式物件
bold = workbook.add_format({'bold':True})

# 5. 向單元格寫入資料
# 5.1 向A1單元格寫入'Hello'
worksheet.write('A1', '員工編號：',bold)
worksheet.write('A2', '諸度股份有限公司 工時紀錄   民國            年            月份     員工：   ',bold)
worksheet.write('A3', '日期')
worksheet.write('B3', '星期')
worksheet.write('C3', '上班')
worksheet.write('D3', '下班')
worksheet.write('E3', '正班時數')
worksheet.write('F3', '加班時數')
worksheet.write('G3', '備註')
worksheet.write('H3', '日期')
worksheet.write('I3', '星期')
worksheet.write('J3', '上班')
worksheet.write('K3', '下班')
worksheet.write('J3', '正班時數')
worksheet.write('M3', '加班時數')
worksheet.write('N3', '備註')

for index, day in enumerate(days):
    worksheet.write(("A" if index < 16 else "H") + str(index % 16 + 4), day)

# 5.7 關閉並儲存文件
workbook.close()