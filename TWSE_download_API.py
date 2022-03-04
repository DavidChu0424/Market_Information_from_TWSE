from logging import root
import requests as req
from openpyxl import Workbook
import time
import pandas as pd


def twsestock(stock, year): 
    # stock = int(input("請輸入股票編號 : "))
    # date = int(input("請輸入歷史資料開始的日期(格式:20100101) : "))

    months = ['01','02','03','04','05','06','07','08','09','10','11','12']

    # wb = Workbook()
    # ws = wb.active

    title = ['日期','成交股數','成交金額','開盤價','最高價','最低價','收盤價','漲跌價差','成交筆數']
    # ws.append(title)

    header = {'user-agent' : 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537.36'}

    totallist = []

    
    for i in range(100):
        try:
            for month in months:
                url = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=' + str(year) + str(month) + '01&' + 'stockNo='+ str(stock) + '&_=1645958064252'
                # print(url)
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
                    # ws.append(info)
                
                    totallist.append(info)
                    # time.sleep(5)

            year = year + 1

        except KeyError:
            break         

    # print(root_json['title'])
    # wb.save(str(root_json['title']) + '.xlsx')
    
    infolist = pd.DataFrame(totallist,columns = title)

    # print(infolist)
    return infolist

# twsestock(2330)
