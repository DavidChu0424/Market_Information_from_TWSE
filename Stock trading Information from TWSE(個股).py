from logging import root
import requests as req
from openpyxl import Workbook

stock = int(input("請輸入股票編號 : "))
date = int(input("請輸入日期(格式:20220101) : "))

wb = Workbook()
ws = wb.active

title = ['日期','成交股數','成交金額','開盤價','最高價','最低價','收盤價','漲跌價差','成交筆數']
ws.append(title)

header = {'user-agent' : 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537.36'}


url = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=' + str(date) + '&' + 'stockNo='+ str(stock) + '&_=1645958064250'
print(url)
r = req.get(url)
root_json = r.json()


for data in root_json['data']:
    info = []
    info.append(data[0])
    info.append(data[1])
    info.append(data[2])
    info.append(data[3])
    info.append(data[4])
    info.append(data[5])
    info.append(data[6])
    info.append(data[7])
    info.append(data[8])
    ws.append(info)

print(root_json['title'])

wb.save(str(root_json['title']) + '.xlsx')



