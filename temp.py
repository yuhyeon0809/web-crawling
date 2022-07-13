import requests 
from bs4 import BeautifulSoup
import openpyxl as op

search_except = ['권총밴드']

url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page=2&max=50&search_text=%EA%B6%8C%EC%B4%9D&orderby=&cid=024000000000000&depth=1"
res = requests.get(url)
res.raise_for_status()
soup = BeautifulSoup(res.text, "lxml")
items = soup.find_all("li", {"class": "float01 search_goods_list"})
i = 1
for item in items:
    title = item.find('div', {"style": "margin-top:10px;height:38px;"})
    itemName = title.find('a').get("title")  # 상품명
    flag = 0
    if itemName.find("<b>") == -1:  # 상품명에 검색어가 포함되지 않은 항목 제외
        continue
    itemName = itemName.replace("<b>", '')
    itemName = itemName.replace("</b>", '')
    
    if itemName.find("밀크북") != -1:
        continue
    if itemName.find("양장") != -1:
        continue
    flag = 0
    for word in search_except:
        if itemName.find(word) != -1:
            flag = 1
    if flag == 1:
        continue


    print(itemName)