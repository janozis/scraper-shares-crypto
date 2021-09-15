from bs4 import BeautifulSoup
from urllib.request import urlopen
import re
import pygsheets
import sys
import os
from datetime import datetime

# the idea is to scrape site for crypto exchange rates
# and the other site for join-stock companies
def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.get_col(1)))
    return str(len(str_list)+1)

def get_crypto_sheet(sh, name):
    try:
        wks = sh.worksheet('title', 'crypto-mapping')
    except:
        print("create a worksheet called 'crypto-mapping'")
        print("closing program")
        exit()
    vals = wks.get_all_values(include_tailing_empty=False,include_tailing_empty_rows=False)
    for val in vals:
        if val[0] == name:
            return val[1]
    return 'no sheet'

def get_companies_sheet(sh, name):
    try:
        wks = sh.worksheet('title', 'shares-mapping')
    except:
        print("create a worksheet called 'shares-mapping'")
        print("closing program")
        exit()
    vals = wks.get_all_values(include_tailing_empty=False,include_tailing_empty_rows=False)
    for val in vals:
        if val[0] == name:
            return val[1]
    return 'no sheet'

def save_new_exchange_rate(c, id, date, rate):
    try:
        sh = c.open_by_key(id)
        wks = sh.sheet1
        next_row = next_available_row(wks)
        wks.cell("A{}".format(next_row)).set_value(date)
        wks.cell("B{}".format(next_row)).set_value(rate)
    except:
        return 0

def get_crypto_exchange_rate(name):
    url = "https://coinmarketcap.com/currencies/"+name+"/"
    page = ""
    price_value = '0'
    try:
        page = urlopen(url)
        html = page.read().decode("utf-8")
        soup = BeautifulSoup(html, "html.parser")
        price_value = soup.find("div", class_="priceValue").text
    except:
        print("wrong name of the cryptocurrency")
        return 0
    if price_value != '0':
        price_value = price_value.replace(",",".")
        price_value = re.findall(r'\d+\.\d+', price_value)[0]
    return price_value

def get_company_exchange_rate(name):
    url = "https://www.bankier.pl/inwestowanie/profile/quote.html?symbol="+name+"/"
    page = ""
    price_value = 0
    try:
        page = urlopen(url)
        html = page.read().decode("utf-8")
        soup = BeautifulSoup(html, "html.parser")
        price_value = soup.find("div", class_="profilLast").text
    except:
        print("wrong name of the company")
        return 0
    price_value = price_value.replace(",",".")
    price_value = re.findall(r'\d+\.\d+', price_value)[0]
    return price_value

# update gsheet companies
def update_sheet_companies(sh,date):
    try:
        wks = sh.worksheet('title', 'shares exchange rates')
    except:
        print("create a worksheet called 'shares exchange rates'")
        print("closing program")
        exit()
    shares = wks.get_values_batch(['B:B'])
    prices = []
    for share in shares[0]:
        if len(share) >= 1:
            prices.append(get_company_exchange_rate(share[0]))
    i = 1
    for price in prices:
        if price != 0:
            cell = 'C' + str(i)
            price = price.replace('.', ',')
            wks.update_value(cell, price)
            sheet_id = get_companies_sheet(sh, shares[0][i-1][0])
            save_new_exchange_rate(c, sheet_id, date, price)
        i += 1

def update_sheet_crypto(sh,date):
    try:
        wks = sh.worksheet('title', 'crypto exchange rates')
    except:
        print("create a worksheet called 'crypto exchange rates'")
        print("closing program")
        exit()
    shares = wks.get_values_batch(['B:B'])
    prices = []
    for share in shares[0]:
        if len(share) >= 1:
            prices.append(get_crypto_exchange_rate(share[0]))
    i = 1
    for price in prices:
        if price != 0:
            cell = 'C' + str(i)
            price = price.replace('.', ',')
            wks.update_value(cell, price)
            sheet_id = get_crypto_sheet(sh, shares[0][i-1][0])
            save_new_exchange_rate(c, sheet_id, date, price)
        i += 1


if __name__ == '__main__':
    type_of_check = ""
    name = ""
    c = pygsheets.authorize(service_file=os.path.dirname(os.path.abspath(__file__)) + '/client_secret.json')
    key = ''
    if len(sys.argv) < 2:
        print("Please add id of gsheet in arguments e.g.")
        print("python3 main.py id_of_gsheet")
        print("closing program")
        exit()
    else:
        key = sys.argv[1]
    try:
        sh = c.open_by_key(key)
    except:
        print("Wrong id of gsheet")
        print("closing program")
        exit()
    now = datetime.now()
    date = now.strftime("%Y-%m-%d %H:%M:%S")
    print("Start")
    print("Update gsheet with companies exchange rates")
    update_sheet_companies(sh, date)
    print("Update gsheet with crypto exchange rates")
    update_sheet_crypto(sh, date)
    print("END")



