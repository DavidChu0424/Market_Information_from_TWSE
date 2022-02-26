from logging import root
import requests as req
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

title = ['成交統計','成交金額(元)','成交股數(股)',"成交筆數"]
ws.append(title)

header = {'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36'}

url = 'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date=&type=&_=1645859100972'
print(url)
r = req.get(url)
root_json = r.json()


for data in root_json['data7']:
    info = []
    info.append(data[0])
    info.append(data[1])
    info.append(data[2])
    info.append(data[3])

    ws.append(info)

print(root_json['subtitle7'])

wb.save(str(root_json['subtitle7']) + '.xlsx')



