#导入bs4模块
from bs4 import BeautifulSoup
import xlrd  # xlrd版本为1.2.0
import urllib
import requests
Skin = [".png", "_2.png", "_skin1.png", "_skin2.png"]

excel_path = "./Source/ark2.xlsx"  # 数据表位置
ExcelBlockLine = 3  # 表前空行数
excel = xlrd.open_workbook(excel_path, encoding_override="utf-8")
AllSheets = excel.sheets()  # 获取所有页对象列表
UseSheet = AllSheets[0]  # 储存第一页对象
LineMax = UseSheet.nrows  # 数据行数-1
for NowLine in range(ExcelBlockLine, LineMax):
    Charactor = UseSheet.cell_value(NowLine, 0)  # 读取角色名(第0+1列)
    print(NowLine)
    for i in range(3):
        url = "https://prts.wiki/w/文件:头像_" + str(Charactor) + Skin[i]
        bianma = urllib.parse.quote(url,safe = "/?:@#$=-+&",encoding="utf-8")
        myURL = requests.get(bianma)
        soup = BeautifulSoup(myURL.text, 'html.parser')
        outmeta = soup.head.find_all('meta',property="og:image")
        out = outmeta[0]['content']
        print(out)
        res = requests.get(out)
        pathphoto = './crt/头像_' + str(Charactor) + Skin[i]
        photo = open(pathphoto, 'wb')
        photo.write(res.content)
        photo.close()

print("over")