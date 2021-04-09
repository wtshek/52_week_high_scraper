import requests
from bs4 import BeautifulSoup
import re
# excel import
from xlutils.copy import copy
from xlrd import open_workbook

# other
from datetime import date

# setup excel
file_path = "C:/Users/tungt/OneDrive/Desktop/stock.xls"

rb = open_workbook(file_path)
sheet_name = "52-week high"
wb = copy(rb)
start_row = rb.sheet_by_name(sheet_name).nrows
w_sheet = wb.get_sheet(sheet_name)


# other info
date = date.today().strftime('%Y-%m-%d')

# scraping info
headers = {
    'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:81.0) Gecko/20100101 Firefox/81.0',
    'X-Requested-With': 'XMLHttpRequest',
}

week_high_url = "https://www.investing.com/equities/52-week-high"
week_high_referer = "https://www.investing.com/equities/52-week-high"
high_headers = {**headers, "Referer": week_high_referer}

page = requests.get(week_high_referer, headers=headers)
soup = BeautifulSoup(page.content, 'html.parser')


tickers = soup.find_all(class_="left bold plusIconTd elp")


equity_base_url = "https://www.investing.com"


col_date = 0
col_ticker = 1
col_ind = 2
col_sect = 3


def write_sheet(ticker, ind, sect):
    global start_row
    w_sheet.write(start_row, col_date, date)
    w_sheet.write(start_row, col_ticker, ticker)
    w_sheet.write(start_row, col_ind, ind)
    w_sheet.write(start_row, col_sect, sect)
    start_row += 1


for ticker in tickers:
    a = ticker.find_all("a", href=True)
    for ele in ticker.select("td[class*='turnover']"):
        vol = ele.get_text()
        if "K" in vol:
            vol_num = int(vol[:len(vol)-1])
            if(vol_num < 500):
                continue
    url = equity_base_url + a[0]["href"]+"-company-profile"
    headers = {**headers, "Referer": url}
    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')

    # get info
    ticker_tag = soup.find(class_="float_lang_base_1 relativeAttr")
    try:
        ticker = ticker_tag.get_text()
    except:
        pass
    header = soup.find(class_="companyProfileHeader")

    industry = ""
    sector = ""

    print(ticker)

    try:
        for child in header.find_all('div'):
            child_text = child.get_text()
            if "Industry" in child_text:
                industry = child.find("a").get_text()
            elif "Sector" in child_text:
                sector = child.find("a").get_text()
            else:
                break
    except:
        pass

    # write sheet
    write_sheet(ticker, industry, sector)


wb.save(file_path)

print("finished")
