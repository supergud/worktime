# !/usr/bin/python
# coding:utf-8
# xlsxwriter的基本用法
import xlsxwriter
import datetime, calendar
import sys

def weekday(x):
    return {
        0: "週一",
        1: "週二",
        2: "週三",
        3: "週四",
        4: "週五",
        5: "週六",
        6: "週日"
    }[x]

year = 2018
month = 1
num_days = calendar.monthrange(year, month)[1]
days = [{'day' : str(month) + '/' + str(day), 'weekday' : (datetime.datetime(year, month, day).weekday())} for day in range(1, num_days + 1)]

# 1. 建立一個Excel文件
workbook = xlsxwriter.Workbook('demo1.xlsx')

# 2. 建立一個工作表sheet物件
worksheet = workbook.add_worksheet()

# 4. 定義一個加粗的格式物件
bold = workbook.add_format({'bold':True})
bg_color = workbook.add_format({'bg_color': '#C0C0C0'})

# 5. 向單元格寫入資料
# 5.1 向A1單元格寫入'Hello'
worksheet.write('A1', '員工編號：',bold)
worksheet.write('A1', '員工編號：           諸度股份有限公司 工時紀錄   民國 ' + str(year - 1911) + ' 年 ' + str(month) + ' 月份     員工：       ', bold)
worksheet.write('A2', '日期', bg_color)
worksheet.write('B2', '星期', bg_color)
worksheet.write('C2', '上班', bg_color)
worksheet.write('D2', '下班', bg_color)
worksheet.write('E2', '正班時數', bg_color)
worksheet.write('F2', '加班時數', bg_color)
worksheet.write('G2', '備註', bg_color)
worksheet.write('H2', '日期', bg_color)
worksheet.write('I2', '星期', bg_color)
worksheet.write('J2', '上班', bg_color)
worksheet.write('K2', '下班', bg_color)
worksheet.write('L2', '正班時數', bg_color)
worksheet.write('M2', '加班時數', bg_color)
worksheet.write('N2', '備註', bg_color)

for index, day in enumerate(days):
    worksheet.write(("A" if index < 16 else "H") + str(index % 16 + 3), day['day'])
    worksheet.write(("B" if index < 16 else "H") + str(index % 16 + 3), weekday(day['weekday']))

worksheet.write('A20', "※每日工作9小時，中午休息一個小時，共為8小時。\n\r※延長工作時數：每日不得超過12小時，每月不得超過46小時。\n\r※出勤紀錄，應逐日記載勞工出勤情形至分鐘為止。依據勞動基準法第 30條規定，應保存五年。")   
worksheet.write('K20', '簽名')

worksheet.set_default_row(25)
worksheet.set_row(0, 35)
worksheet.set_row(1, 20)
worksheet.set_row(18, 50)

# 5.7 關閉並儲存文件
workbook.close()