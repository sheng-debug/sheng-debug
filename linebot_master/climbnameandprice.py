from bs4 import BeautifulSoup
from urllib.request import urlopen, urlretrieve
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import pandas as pd

df = pd.DataFrame(columns=["商品照片", "商品名稱", "商品價格"])
import warnings

# 忽略掉對ignore的warning
warnings.filterwarnings('ignore')

from styleframe import StyleFrame
import glob
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing import image
import os
import requests

url = 'https://www.maccosmetics.com.tw/products/13854'

res = requests.get(url)
soup = BeautifulSoup(res.text, 'html.parser')

for i in soup.find_all("div", class_="product__detail"):  # => find_all 必轉出 list

    name = i.find("h3", class_="product__subline")
    names = list(name.stripped_strings) #把爬出來的東西切成只有文字，這行可要可不要
    price = i.find("span", class_="product__price--standard")
    prices = list(price.stripped_strings) #把爬出來的東西切成只有文字，這行可要可不要
    #img = i.find("img", class_="product__sku-image")

    #fname = "goods/" + img["src"].split("/")[-1]
    #urlretrieve(img["src"], fname)

    # 準備Series 以及 append進DataFrame。值會放到相對印的column
    s = pd.Series([name.text, price.text],
                  index=["商品名稱", "商品價格"])
    # 因為 Series 沒有橫列的標籤, 所以加進去的時候一定要 ignore_index=True
    df = df.append(s, ignore_index=True)
    df.to_excel("goods.xlsx", encoding="utf-8", index=False)
sf = StyleFrame(df)
# 設定欄寬
sf.set_column_width_dict(col_width_dict={
        ("商品照片"): 25.5,
    ("商品名稱", "商品價格"): 20,
        # ("介紹網址"): 65.5
})

# 設定列高
all_rows = sf.row_indexes
sf.set_row_height_dict(row_height_dict={
    all_rows[1:]: 120
})
# 存成excel檔
sf.to_excel('goods.xlsx',
            sheet_name='Sheet1',  # Create sheet
            right_to_left=False,  # False 所以sheet放置是從左到右 left-to-right
            columns_and_rows_to_freeze='A1',  # 資料從A1整個貼上
            row_to_add_filters=0).save()  # 不要忘記要save
col = 0
wb = load_workbook('goods.xlsx')
ws = wb.worksheets[0]

searchedfiles = sorted(glob.glob("goods/*.jpg"), key=os.path.getmtime)
for fn in searchedfiles:
    img = openpyxl.drawing.image.Image(fn)  # create image instances
    c = str(col + 2)
    ws.add_image(img, 'A' + c)
    col = col + 1
wb.save('goods.xlsx')