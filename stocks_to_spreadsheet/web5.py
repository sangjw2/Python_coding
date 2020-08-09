import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import string
import time


class GetSource_MakeSheet():
    real_pages = "https://finance.naver.com/sise/sise_market_sum.nhn?sosok="
    name_number = ['현재가', '전일비', '등락률', '액면가', '시가총액', '상장주식', '외국인비율', '거래량', 'PER', 'ROE']
    a = []
    b = {}
    c = []
    def __init__(self, cos, pages):
        if cos == 1:
            self.html = requests.get(self.real_pages+"0")  # cospi
        elif cos == 2:
            self.html = requests.get(self.real_pages+"1")  # cosdaq
        else:
            print("invalid value")
            return

    def parsing_pages(self, pages):
        # Get value from Internet
        self.soup = BeautifulSoup(self.html.content, "html.parser")
        self.num = self.soup.select('td.number')
        self.corp = self.soup.find_all('a', attrs={'class': 'tltle'})
        for h in self.corp:
            self.c.append(str(h.get_text()))
        

def GetSource_MakeSheet(cos, pages):
    real_pages += "&page="
    for wp in pages:
        real_pages += "1"
        
        how_first = 0
        how = 10
        
        for x in num:
            a.append(str((x.get_text()).strip()))

        for j in range(50):
            for i in a[how-10:how]:
                b[c[j]+' '+name_number[how_first]] = i
                how_first += 1
            how_first = 0
            how += 10

        # SpreadSheet
        wb = Workbook()
        ws = wb.active

        ws.title = "Main Sheet"

        cell = ws['A1']
        cell.value = "종목명"

        for m in c:  # 종목명 삽입
            cell = ws['A'+str(50*pages+(c.index(m)+2))]  # 2345
            cell.value = m

        alphabet = list(string.ascii_uppercase)

        for n in name_number:  # name_number 삽입
            s = name_number.index(n)
            cell = ws[alphabet[s+1]+'1']
            cell.value = n

        for i in name_number:  # insert every value on cell
            for j in c:
                s = name_number.index(i)
                cell = ws[alphabet[s+1]+str(c.index(j)+2)]
                cell.value = b[j+' '+i]
            ws.column_dimensions[alphabet[name_number.index(i)]].auto_size = True  # 열 너비 자동 맞춤


    today = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    wb.save(today+".xlsx")

which_data = input("Data Source를 선택해주십시오 (1 : 코스피,2 : 코스닥): ")
howmany_page = input("몇 페이지까지 확인하시겠습니까?(1페이지 당 50종목): ")
GetSource_MakeSheet(which_data,howmany_page)
