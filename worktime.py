# !/usr/bin/python
# coding:utf-8
# xlsxwriter的基本用法
import xlsxwriter

# 1. 建立一個Excel文件
workbook = xlsxwriter.Workbook('demo1.xlsx')

# 2. 建立一個工作表sheet物件
worksheet = workbook.add_worksheet()

# 3. 設定第一列（A）寬度為20畫素
worksheet.set_column('A:A',20)

# 4. 定義一個加粗的格式物件
bold = workbook.add_format({'bold':True})

# 5. 向單元格寫入資料
# 5.1 向A1單元格寫入'Hello'
worksheet.write('A1','Hello')
# 5.2 向A2單元格寫入'World'並使用bold加粗格式
worksheet.write('A2','World',bold)
# 5.3 向B2單元格寫入中文並使用加粗格式
worksheet.write('B2',u'中文字符',bold)

# 5.4 用行列表示法（行列索引都從0開始）向第2行、第0列（即A3單元格）和第3行、第0列（即A4單元格）寫入數字
worksheet.write(2,0,10)
worksheet.write(3,0,20)

# 5.5 求A3、A4單元格的和並寫入A5單元格，由此可見可以直接使用公式
worksheet.write(4,0,'=SUM(A3:A4)')

# 5.6 在B5單元格插入圖片
worksheet.insert_image('B5','./demo.png')

# 5.7 關閉並儲存文件
workbook.close()