import csv
import requests 
from bs4 import BeautifulSoup
from openpyxl import Workbook
import pandas as pd
import urllib.request

# 물품리스트(엑셀파일) 읽어오기
def read_list(list_xlsx):
    df = pd.read_excel(list_xlsx, usecols = "A")
    item_list = df.values.tolist()
    return item_list

# 크롤링한 텍스트 데이터를 csv파일로 출력
def write_csv(item_list, imgPath, writer):

    for group in item_list:
        search = str(group)
        search = search[2:-2]
        print(search)

        for page in range(99, 122): 
            print(page)
            # 크롤링 대상 url
            url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page={}&max=50&search_text={}&orderby=daiso_ranking1".format(page, search)

            res = requests.get(url)
            res.raise_for_status()

            soup = BeautifulSoup(res.text, "lxml")
            items = soup.find_all("li", {"class": "float01 search_goods_list"})
            # print(items)

            i = 1
            for item in items:
                title = item.find('div', {"style": "margin-top:10px;height:38px;"})
                itemName = title.find('a').get("title")
                if itemName.find("<b>") == -1:  # 상품명에 검색어가 포함되지 않은 항목 제외
                    continue
                itemName = itemName.replace("<b>", '')
                itemName = itemName.replace("</b>", '')

                price = item.find('div', {"style": "margin-top:12px;"})
                itemPrice = price.find('strong').text
                itemPrice = itemPrice.replace("원", '')
                itemPrice = itemPrice.replace(",", '')

                img = item.find('div', {"class": "goods_line_img"})
                imgUrl = img.find('img').get('src')
                imgName = "{}{}.jpg".format(search, (page-1)*50+i)

                # 물품 이미지 다운로드
                urllib.request.urlretrieve(imgUrl, imgPath+imgName)

                itemId = item.find('a').get('href')
                itemId = itemId[24:34]
                itemUrl = "https://www.daisomall.co.kr/shop/goods_view.php?id={}&depth=1&search_text={}".format(itemId, search)

                itemRes = requests.get(itemUrl)
                itemRes.raise_for_status()
                soupItem = BeautifulSoup(itemRes.text, "lxml")

                categories = []
                errors = {'search':'itemUrl'}

                try:
                    itemNum = soupItem.find('td', {"class": "color_63 line_h160"}).find('strong').text
                except:
                    print("today deal error!")
                    errors["today deal error"] = itemUrl
                    continue
                
                for category in soupItem.find_all('option', {"selected": ""}):
                    if str(category).find("selected") == -1:
                        continue
                    category = category.get_text()
                    categories.append(category)

                try:
                    data = [None, search, (page-1)*50+i, itemNum, categories[0]+">"+categories[1]+">"+categories[2], itemName, imgUrl, itemPrice]
                except:
                    print("catetory error!")
                    errors["catetory error"] = itemUrl
                    continue

                writer.writerow(data)
                i+=1

    print("------error url------")
    print("총 " + str(len(errors)) + "개")
    for error in errors:
        print(error +"\n")

# csv파일을 엑셀파일로 변환
def csvtoxlsx(filename_csv):
    wb = Workbook()
    ws = wb.active
    with open(filename_csv, 'r', encoding='utf8') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save('item_test.xlsx')


# main 함수
if __name__ == "__main__":

    filename_csv = "item_info.csv"  # 결과를 저장할 csv 파일 이름
    imgPath = "item_img/"           # 이미지 파일이 저장될 경로

    f = open(filename_csv, "a", encoding="utf-8-sig", newline="")
    writer = csv.writer(f)


    # 컬럼 이름 지정
    columns_name = ["물품분류", "물품종", "순번", "상품번호", "카테고리", "상품명", "상품사진", "가격"] 
    writer.writerow(columns_name)

    item_list = read_list("item_list.xlsx")
    write_csv(item_list, imgPath, writer)
    csvtoxlsx(filename_csv)
