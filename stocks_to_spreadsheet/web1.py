import requests
from bs4 import BeautifulSoup

html = requests.get("https://finance.naver.com/sise/sise_market_sum.nhn")
soup = BeautifulSoup(html.content, "html.parser")
num = soup.select('td.number')
#num = soup.find_all('td', attrs={'class': 'number'})
corp = soup.find_all('a', attrs={'class': 'tltle'})
# for n in corp():
#     print(n.get_text())
a = []
# for x in range(10):
#     a.append((num[x].get_text()))
# print(a, type(num))
for x in num:
    a.append(str((x.get_text()).strip()))
print(a)
