import requests 
import pandas as pd
from bs4 import BeautifulSoup
import openpyxl as op
import csv

def read_list(list_xlsx):
    df = pd.read_excel(list_xlsx, usecols = ['검색어'], skiprows=[1])
    item_list = df.values.tolist()
    return item_list

def write_csv(item_list, writer):

    for group in item_list:
        search = str(group)
        search = search[2:-2]
        print(search)

        url = "https://www.daisomall.co.kr/shop/search.php?nset=1&max=50&search_text={}&orderby=daiso_ranking1".format(search)
        
        res = requests.get(url)
        res.raise_for_status()
        
        soup = BeautifulSoup(res.text, "lxml")
        try:
            max_item = int(soup.find("span", {"class": "font_normal size_16"}).text)
        except:
            max_item = 0

        categories = []
        for categoryBox in soup.find_all("li", {"class":"float01"}):
            try:
                category = categoryBox.find('a').text
                print(category)
                categories.append(category)
            except:
                break

        data = [None, None, None, None, None, "총 상품 " + str(max_item) + "개"]
        for i in range(0, len(categories)):
            data.append(categories[i])

        writer.writerow(data)

def csvtoxlsx(filename_csv, filename_xlsx):
    wb = op.Workbook()
    ws = wb.active
    with open(filename_csv, 'r', encoding='utf8') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(filename_xlsx)

    wb = op.load_workbook(filename_xlsx) 
    ws = wb.active
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 15
    
    for i in range(0, 10):
        ws.column_dimensions[chr(67+i)].width = 20

    wb.save(filename_xlsx)
    

if __name__ == "__main__":

    filename_csv = "item_category_list.csv"  # 결과를 저장할 csv 파일 이름
    filename_xlsx = "item_category_list.xlsx"

    # f = open(filename_csv, "w", encoding="utf-8-sig", newline="")
    # writer = csv.writer(f)


    # item_list = read_list("item_list.xlsx")
    # write_csv(item_list, writer)
    # f.close()
    csvtoxlsx(filename_csv, filename_xlsx)
        