import requests
from bs4 import BeautifulSoup

html = requests.get("https://finance.naver.com/sise/sise_market_sum.nhn")
soup = BeautifulSoup(html.content, "html.parser")
num = soup.select('td.number')
corp = soup.find_all('a', attrs={'class': 'tltle'})
c = []
for h in corp:
    c.append(str(h.get_text()))
a = []
b = {}
how_first = 0
how = 10
name_number = ['현재가', '전일비', '등락률', '액면가',
               '시가총액', '상장주식', '외국인비율', '거래량', 'PER', 'ROE']
for x in num:
    a.append(str((x.get_text()).strip()))

for j in range(50):
    for i in a[how-10:how]:
        b[c[j]+' '+name_number[how_first]] = i
        how_first += 1
    how_first = 0
    how += 10
print(b)
