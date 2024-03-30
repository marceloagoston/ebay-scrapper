from bs4 import BeautifulSoup
import requests
import openpyxl

from datetime import datetime

def get_page_elements(p):
    prod = p.find_all("div", class_="s-item__wrapper clearfix")
    data = []
    for p in prod:
       
        sold = p.find_all("div", class_="s-item__caption-section")
        title = p.find_all("div", class_="s-item__title")
        price = p.find_all("span", class_="s-item__price")
        link = title[0].parent
                
        export_title = title[0].findChildren()[0].text
        sold_price = price[0].findChildren()[0].string if price[0].findChildren() else None

        date_obj = None
        if sold:
            date_obj = sold[0].findChildren()[0].findChildren()[0].string.replace("Sold", "").strip()

        if not "Shop on eBay" in export_title:
            data.append([export_title, link['href'], sold_price, date_obj])

    return data


# Create a new Workbook
workbook = openpyxl.Workbook()

# Select the active worksheet
worksheet = workbook.active

url_input = input("Enter a URL to search in ebay:").strip()

url = url_input

# handle pagination
current_page = 1
page = f"&_pgn={current_page}"
if "&_pgn=1" in url:
    url = url.replace("&_pgn=1", "")

data = []

try:
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')

    pages = 0
    if soup.find_all("ol", class_="pagination__items"):
        pages = int(len(soup.find_all("ol", class_="pagination__items")[0].findChildren()) / 2)
    
    if soup and pages == 0:
        data += get_page_elements(soup)

    elif soup and pages > 0:
        for pg in range(1,pages+1):
            page = requests.get(f"{url}&_pgn={pg}")
            soup = BeautifulSoup(page.content, 'html.parser')
            data += get_page_elements(soup)

    else:
        print('Nothing in this URL or bad URL')

except:
    print('Error: Bad URL')


worksheet.append(["Product Name", "URL", "Sold Price", "Date"])

for row_data in data:
    worksheet.append(row_data)

current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

workbook.save(f"output_{current_datetime}.xlsx")
