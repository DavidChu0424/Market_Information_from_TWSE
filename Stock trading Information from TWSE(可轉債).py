from bs4 import BeautifulSoup
from logging import root
from matplotlib.pyplot import title
import requests as req
from openpyxl import Workbook
import pandas as pd


url = 'https://www.tpex.org.tw/web/bond/publish/convertible_bond_search/memo.php?l=zh-tw'
response = req.get(url)
html = response.text
soup = BeautifulSoup(html, 'html.parser')

table = soup.find('table', class_ = 'table table-striped table-bordered')

title = ["債卷代號", "債卷簡稱", "發行人", "發行日期", "到期日期", "年期", "發行總面額", "發行資料"]

list = []
a_tags = soup.find_all('td')
for i in range(5):
    for tag in a_tags:
        tagresult = tag.string
        list.append(tagresult)


n = 8
output=[list[i:i + n] for i in range(0, len(list), n)]

infolist = pd.DataFrame(output,columns = title)

infolist['年期'] = infolist['年期'].str.replace('å¹´\r\n\t\t\t', "")

infolist.to_csv("發債股.csv", index = False)

print(infolist)