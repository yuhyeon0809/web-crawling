import csv
import requests 
from bs4 import BeautifulSoup

url = "https://www.daisomall.co.kr/event/goods_best.php"

filename = "test_itemname.csv"
f = open(filename, "w", encoding="utf-8-sig", newline="")
writer = csv.writer(f)

columns_name = ["순위", "상품명"] # 컬럼명

writer.writerow(columns_name)

res = requests.get(url)
res.raise_for_status()

soup = BeautifulSoup(res.text, "lxml")
itemBox = soup.find('ul', attrs={"class": "goodsBox float01"})
items = itemBox.find_all('a')

i = 1

for item in items:
    title = item.get("title")
    if title==None:
        continue
    print(f"{str(i)}: {title}")
    data = [str(i), title]
    writer.writerow(data)
    i += 1