import csv
import requests 
from bs4 import BeautifulSoup
import openpyxl as op
import urllib.request
import pandas as pd

search = "가위"

def test(item_list_xlsx, imgPath, writer):
    url = "https://www.daisomall.co.kr/shop/search.php?search_text=%EA%B0%80%EC%9C%84&x=18&y=23"
    res = requests.get(url)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "lxml")
    items = soup.find_all("li", {"class": "float01 search_goods_list"})

    num = 1
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

        print(itemName)
        price = item.find('div', {"style": "margin-top:12px;"})
        itemPrice = price.find('strong').text
        itemPrice = itemPrice.replace("원", '')
        itemPrice = itemPrice.replace(",", '')

        img = item.find('div', {"class": "goods_line_img"})
        imgUrl = img.find('img').get('src')
        imgName = "{}{}.jpg".format(search, num)

        # 물품 이미지 다운로드
        urllib.request.urlretrieve(imgUrl, imgName)
        itemId = item.find('a').get('href')
        itemId = itemId[24:34]
        itemUrl = "https://www.daisomall.co.kr/shop/goods_view.php?id={}&depth=1&search_text={}".format(itemId, search)

        itemRes = requests.get(itemUrl)
        itemRes.raise_for_status()
        soupItem = BeautifulSoup(itemRes.text, "lxml")
        itemCategories = []
        try:
            itemNum = soupItem.find('td', {"class": "color_63 line_h160"}).find('strong').text
        except:
            print("today deal error!")
            continue
        
        for itemCategory in soupItem.find_all('option', {"selected": ""}):
            if str(itemCategory).find("selected") == -1:
                continue
            itemCategory = itemCategory.get_text()
            itemCategories.append(itemCategory)
        try:
            data = ["가위", search, num, itemNum, itemCategories[0]+">"+itemCategories[1]+">"+itemCategories[2], itemName, imgUrl, itemPrice]
        except:
            print("category error!")
            continue
        writer.writerow(data)
        num+=1

# csv파일을 엑셀파일로 변환
def csvtoxlsx(filename_csv, filename_xlsx):
    wb = op.Workbook()
    ws = wb.active
    with open(filename_csv, 'r', encoding='utf8') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(filename_xlsx)

# 중복 행 제거
def drop_duplicates(filename_xlsx):
    df = pd.read_excel(filename_xlsx)  # 엑셀파일 읽어오기
    df = df.drop_duplicates(['상품번호'])  # 중복 행 제거
    df.to_excel(filename_xlsx)  # 중복 제거된 파일 저장

if __name__ == "__main__":
    
    item_list_xlsx = "item_category_list.xlsx"  # 읽어올 물품 리스트
    filename_csv = "item_category_test.csv"          # 결과를 저장할 csv 파일 이름
    filename_xlsx = "item_category_result_test.xlsx"        # 결과를 저장할 xlsx 파일 이름
    imgPath = "/"                       # 이미지 파일이 저장될 경로
    
    # f = open(filename_csv, "a", encoding="utf-8-sig", newline="")
    # writer = csv.writer(f)

    # # 컬럼 이름 지정
    # columns_name = ["물품분류", "물품종", "순번", "상품번호", "카테고리", "상품명", "상품사진", "가격"] 
    # writer.writerow(columns_name)

    # test(item_list_xlsx, imgPath, writer)
    # f.close()
    # csvtoxlsx(filename_csv, filename_xlsx)
    drop_duplicates(filename_xlsx)