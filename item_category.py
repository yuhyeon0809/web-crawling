import csv
import requests 
from bs4 import BeautifulSoup
import openpyxl as op
import urllib.request

toCid = {'수납/정리': '019000000000000',
        '주방/욕실/청소': '020000000000000',
        '가구/인테리어': '022000000000000',
        '사무/문구/디지털': '021000000000000',
        '가전/레져/식품': '023000000000000',
        '키즈/뷰티/패션잡화': '024000000000000',
        '다이소 매장상품': '010000000000000',
        '포장재 전문관': '027000000000000'}

def write_csv(item_list_xlsx, imgPath, writer):

    wb = op.load_workbook(item_list_xlsx) # 엑셀파일 열기
    ws = wb.active
 
    row_max = ws.max_row # 최대행값 저장

    for r in range(2, row_max+1):  # 2행부터 마지막행까지 반복
        categories = []
        type = str(ws.cell(row=r, column=1).value)  # 물품분류
        search = str(ws.cell(row=r, column=2).value)  # 검색어
        if search == 'None':
            continue
        search_include = str(ws.cell(row=r, column=3).value).split(', ')    # 포함검색어
        search_except = str(ws.cell(row=r, column=4).value).split(', ')     # 제외검색어
        # max_item = int(str(ws.cell(row=r, column=6).value)[5:-1])           # 총 물품 개수

        for c in range(7, ws.max_column+1):  # 카테고리명 리스트에 저장
            category = str(ws.cell(r, c).value)
            print(category)
            if category == 'None':
                continue
            categories.append(category)

        num = 1
        for category in categories:
            index1 = category.find('(')
            index2 = category.find(')')
            itemNum = int(category[index1+1:index2].replace(',', ''))
            categoryPage = int(itemNum/50) + 1
            category = category[:index1]
            try:
                cid = toCid[category]
            except:
                print(category + "---unknown category!")
                continue

            for page in range(1, categoryPage+1): 
                print(page)
                url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page={}&max=50&search_text={}&orderby=daiso_ranking1&cid={}&depth=1".format(page, search, cid)
                res = requests.get(url)
                res.raise_for_status()
                soup = BeautifulSoup(res.text, "lxml")

                items = soup.find_all("li", {"class": "float01 search_goods_list"})

                for item in items:
                    title = item.find('div', {"style": "margin-top:10px;height:38px;"})

                    itemName = title.find('a').get("title")  # 상품명
                    flag = 0
                    if itemName.find("<b>") == -1:  # 상품명에 검색어가 포함되지 않은 항목 제외
                        for word in search_include:
                            if itemName.find(word) != -1:
                                flag = 1
                        if flag == 0:
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

                    price = item.find('div', {"style": "margin-top:12px;"})
                    itemPrice = price.find('strong').text
                    itemPrice = itemPrice.replace("원", '')
                    itemPrice = itemPrice.replace(",", '')

                    img = item.find('div', {"class": "goods_line_img"})
                    imgUrl = img.find('img').get('src')
                    imgName = "{}{}.jpg".format(search, num)

                    # 물품 이미지 다운로드
                    urllib.request.urlretrieve(imgUrl, imgPath+imgName)

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
                        data = [type, search, num, itemNum, itemCategories[0]+">"+itemCategories[1]+">"+itemCategories[2], itemName, imgUrl, itemPrice]
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


# main 함수
if __name__ == "__main__":
    
    item_list_xlsx = "item_category_list.xlsx"  # 읽어올 물품 리스트
    filename_csv = "item_category.csv"          # 결과를 저장할 csv 파일 이름
    filename_xlsx = "item_category.xlsx"        # 결과를 저장할 xlsx 파일 이름
    imgPath = "item_img/"                       # 이미지 파일이 저장될 경로
    
    f = open(filename_csv, "w", encoding="utf-8-sig", newline="")
    writer = csv.writer(f)

    # 컬럼 이름 지정
    columns_name = ["물품분류", "물품종", "순번", "상품번호", "카테고리", "상품명", "상품사진", "가격"] 
    writer.writerow(columns_name)

    write_csv(item_list_xlsx, imgPath, writer)
    f.close()
    csvtoxlsx(filename_csv, filename_xlsx)