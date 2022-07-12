import csv
import requests 
from bs4 import BeautifulSoup
from openpyxl import Workbook
import pandas as pd
import urllib.request

def read_list(list_xlsx):
    df = pd.read_excel(list_xlsx, usecols = "A")
    item_list = df.values.tolist()
    return item_list

def write_csv(item_list, writer):

    for group in item_list:
        search = str(group)
        search = search[2:-2]
        print(search)

        for page in range(1, 10):
            url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page={}&max=50&search_text={}&orderby=daiso_ranking1".format(page, search)

            res = requests.get(url)
            res.raise_for_status()

            soup = BeautifulSoup(res.text, "lxml")
            items = soup.find_all("li", {"class": "float01 search_goods_list"})

            i = 1
            for item in items:
                title = item.find('div', {"style": "margin-top:10px;height:38px;"})
                itemName = title.find('a').get("title")
                itemName = itemName.replace("<b>", '')
                itemName = itemName.replace("</b>", '')

                price = item.find('div', {"style": "margin-top:12px;"})
                itemPrice = price.find('strong').text
                itemPrice = itemPrice.replace("원", '')
                itemPrice = int(itemPrice.replace(",", ''))

                img = item.find('div', {"class": "goods_line_img"})
                imgUrl = img.find('img').get('src')

                urllib.request.urlretrieve(imgUrl, "{}{}.png".format(item, (page-1)*50+i))

                data = [None, search, (page-1)*50+i, 0, itemName, imgUrl, itemPrice]

                writer.writerow(data)
                i+=1

def csvtoexcel(filename_csv):
    wb = Workbook()
    ws = wb.active
    with open(filename_csv, 'r', encoding='utf8') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save('item_test.xlsx')
    
if __name__ == "__main__":

    filename_csv = "item_.csv"
    f = open(filename_csv, "w", encoding="utf-8-sig", newline="")
    writer = csv.writer(f)

    columns_name = ["물품분류", "물품종", "순번", "물품번호", "카테고리", "상품명", "물품사진", "가격"] 
    writer.writerow(columns_name)

    item_list = ['가위', '칼']
    write_csv(item_list, writer)
    csvtoexcel(filename_csv)