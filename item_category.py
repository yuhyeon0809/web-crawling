import requests 
from bs4 import BeautifulSoup
import openpyxl as op
import urllib.request
import pandas as pd

# cid(카테고리 id) 정의
toCid = {'수납/정리': '019000000000000',
        '주방/욕실/청소': '020000000000000',
        '가구/인테리어': '022000000000000',
        '사무/문구/디지털': '021000000000000',
        '가전/레져/식품': '023000000000000',
        '키즈/뷰티/패션잡화': '024000000000000',
        '다이소 매장상품': '010000000000000',
        '포장재 전문관': '027000000000000'}

def write_data(item_list_xlsx, filename_xlsx, imgPath, sheet_name):

    input_wb = op.load_workbook(item_list_xlsx)  # 입력을 읽어올 엑셀파일
    input_ws = input_wb[sheet_name]
    row_max = input_ws.max_row # 최대행값 저장

    output_wb = op.load_workbook(filename_xlsx)  # 결과를 저장할 엑셀파일
    output_ws = output_wb[sheet_name]

    for r in range(2, row_max+1):  # 2행부터 마지막행까지 반복
        categories = []
        type = str(input_ws.cell(row=r, column=1).value)  # 물품분류
        if type == 'None':
            type = str(input_ws.cell(row=r-1, column=1).value)
        search = str(input_ws.cell(row=r, column=2).value)  # 검색어
        if search == 'None':
            continue
        search_include = str(input_ws.cell(row=r, column=3).value).split(', ')    # 포함단어
        search_except = str(input_ws.cell(row=r, column=4).value).split(', ')     # 제외단어
        # max_item = int(str(ws.cell(row=r, column=6).value)[5:-1])           # 총 물품 개수

        for c in range(7, input_ws.max_column+1):  # 카테고리명을 리스트에 저장
            category = str(input_ws.cell(r, c).value)
            if category == 'None':
                continue
            categories.append(category)

        print("-------------" + search + "-------------")
        num = 1  # 순번
        for category in categories:  # 해당 검색어의 카테고리 하나씩 탐색
            index1 = category.find('(')
            index2 = category.find(')')
            itemNum = int(category[index1+1:index2].replace(',', ''))
            categoryPage = int(itemNum/50) + 1
            if categoryPage > 150:  # 150 페이지가 넘어가는 경우 150 페이지까지만 탐색
                categoryPage = 150
            category = category[:index1]
            try:
                cid = toCid[category]  # 한글로 된 카테고리명을 url에서 쓰이는 cid로 변환
            except:
                print(category + "---unknown category!")
                continue

            print("-------------" + category + "-------------")
            for page in range(1, categoryPage+1):  # 1페이지부터 마지막 페이지까지 반복
                print(page)

                # 페이지 url에서 http 소스 읽어옴
                url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page={}&max=50&search_text={}&orderby=daiso_ranking1&cid={}&depth=1".format(page, search, cid)
                res = requests.get(url, headers={'User-Agent':'Mozilla/5.0'})
                res.raise_for_status()
                soup = BeautifulSoup(res.text, "lxml")

                # 해당 페이지에 나와 있는 상품을 모두 저장
                items = soup.find_all("li", {"class": "float01 search_goods_list"})

                # 상품을 하나씩 탐색
                for item in items:
                    title = item.find('div', {"style": "margin-top:10px;height:38px;"})
                    itemName = title.find('a').get("title")  # 상품명
                    flag = 0
                    if itemName.find("<b>") == -1:  # 상품명에 검색어가 포함되지 않은 항목 제외
                        for word in search_include: # 검색어가 포함되지 않았으나 '포함단어' 리스트에 있는 단어를 포함한 경우 제외하지 않음
                            if itemName.find(word) != -1:
                                flag = 1
                        if flag == 0:
                            continue
                    
                    itemName = itemName.replace("<b>", '')
                    itemName = itemName.replace("</b>", '')

                    # 모든 상품에 일괄적으로 제외할 단어들
                    if itemName.find("밀크북") != -1:
                        continue
                    if itemName.find("양장") != -1:
                        continue
                    
                    # '제외단어' 리스트에 있는 단어가 하나라도 포함된 경우 제외
                    flag = 0
                    for word in search_except:
                        if itemName.find(word) != -1:
                            flag = 1
                    if flag == 1:
                        continue

                    print(itemName)

                    price = item.find('div', {"style": "margin-top:12px;"})  # 상품가격
                    itemPrice = price.find('strong').text
                    itemPrice = itemPrice.replace("원", '')
                    itemPrice = itemPrice.replace(",", '')

                    img = item.find('div', {"class": "goods_line_img"})  # 상품 이미지 url
                    imgUrl = img.find('img').get('src')
                    imgName = "{}{}.jpg".format(search, num)  # 이미지 파일 이름 형식

                    # 상품 이미지 다운로드
                    urllib.request.urlretrieve(imgUrl, imgPath+imgName)

                    itemId = item.find('a').get('href')  # 상품번호
                    itemId = itemId[24:34]

                    # 상품번호로 해당 상품의 상세페이지 url 접속해 http 소스 읽어옴
                    itemUrl = "https://www.daisomall.co.kr/shop/goods_view.php?id={}&depth=1&search_text={}".format(itemId, search)
                    try:
                        itemRes = requests.get(itemUrl, headers={'User-Agent':'Mozilla/5.0'})
                    except:
                        print("request error!")
                        continue
                    itemRes.raise_for_status()
                    soupItem = BeautifulSoup(itemRes.text, "lxml")

                    try:
                        itemNum = soupItem.find('td', {"class": "color_63 line_h160"}).find('strong').text 
                    except:
                        print("itemNum error!")
                        continue
                    
                    itemCategories = []
                    for itemCategory in soupItem.find_all('option', {"selected": ""}):  # 해당 상품의 카테고리 저장
                        if str(itemCategory).find("selected") == -1:
                            continue
                        itemCategory = itemCategory.get_text()
                        itemCategories.append(itemCategory)

                    try:
                        data = [type, None, search, num, int(itemNum), itemCategories[0]+">"+itemCategories[1]+">"+itemCategories[2], itemName, imgUrl, int(itemPrice), itemUrl]
                    except:
                        print("category error!")
                        continue
                    output_ws.append(data)
                    num+=1
                    
            output_wb.save(filename_xlsx)


# 물품코드 로딩
def load_code(item_list_xlsx, filename_xlsx, sheet_name):
    wb = op.load_workbook(filename_xlsx)
    ws = wb.active
    code_wb = op.load_workbook(item_list_xlsx)
    code_ws = code_wb['물품코드표']
    row_max = ws.max_row
    code_row_max = code_ws.max_row
    
    codeDic = {'item': 'code'}
    for r in range(2, code_row_max+1):
        code = str(code_ws.cell(row=r, column=1).value)
        item = str(code_ws.cell(row=r, column=2).value)
        desc = str(code_ws.cell(row=r, column=3).value)
        if desc.find("삭제") != -1:
            continue
        codeDic[item] = code

    for r in range(2, row_max+1):
        temp = str(ws.cell(row=r, column=1).value)
        try:
            ws.cell(row=r, column=2).value = int(codeDic[temp])
        except:
            print("Key error! --- " + temp)
            continue
        
    wb.save(filename_xlsx)

# 중복 행 제거
def drop_duplicates(filename_xlsx):
    df = pd.read_excel(filename_xlsx, engine='openpyxl')  # 엑셀파일 읽어오기
    df = df.drop_duplicates(subset='상품번호')  # 중복 행 제거
    df.to_excel(filename_xlsx[:-5]+"_dptest.xlsx", index=False)

# main 함수
if __name__ == "__main__":

    sheet_name = '잡화 슈즈 명품'
    item_list_xlsx = "촬영 대상 물품 분류체계_v0.1_권혁진_다이소몰 크롤링 목록_일반물품.xlsx"  # 읽어올 물품 리스트
    filename_xlsx = "촬영 대상 물품 분류체계_v0.1_권혁진_다이소몰 크롤링 결과_잡화 슈즈 명품_텍스트.xlsx"  # 결과를 저장할 xlsx 파일 이름
    imgPath = "item_img/"  # 이미지 파일이 저장될 경로
    columns_name = ["물품분류", "물품코드", "물품종", "순번", "상품번호", "카테고리", "상품명", "상품사진", "가격", "링크"]  # 컬럼명 지정

    # #--- 새 엑셀파일 생성 시
    output_wb = op.Workbook()
    output_ws = output_wb.create_sheet(sheet_name)
    output_ws.append(columns_name)
    output_wb.save(filename_xlsx)

    # #--- 기존 엑셀파일에 추가 시
    # output_wb = op.load_workbook(filename_xlsx)  # 결과를 저장할 엑셀파일
    # output_ws = output_wb.create_sheet(sheet_name)
    # output_ws.append(columns_name)
    # output_wb.save(filename_xlsx)

    write_data(item_list_xlsx, filename_xlsx, imgPath, sheet_name)
    # drop_duplicates(filename_xlsx)
    # load_code(item_list_xlsx, filename_xlsx, sheet_name) # !!!!!! 주석처리 확인 !!!!!!