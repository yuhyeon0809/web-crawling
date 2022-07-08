import requests 
import urllib.request
from bs4 import BeautifulSoup

page = 90
search = "가위"

url = "https://www.daisomall.co.kr/shop/search.php?nset=1&page={}&max=50&search_text={}&orderby=daiso_ranking1".format(page, search)

res = requests.get(url)
res.raise_for_status()
soup = BeautifulSoup(res.text, "lxml")
items = soup.find_all("li", {"class": "float01 search_goods_list"})

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
    imgName = "{}{}.png".format(search, (page-1)*50+i)
     
     # 물품 이미지 다운로드
    # urllib.request.urlretrieve(imgUrl, imgPath+imgName)
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
        errors[search] = itemUrl
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
        errors[search] = itemUrl
        continue
    
    print(data[3]+" "+data[4]+" "+data[5]+" "+data[7])
    i+=1