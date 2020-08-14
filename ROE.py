import requests
import math
import pygsheets
import urllib.request
from bs4 import BeautifulSoup


def get_price(symb):
    url = "https://finance.yahoo.com/quote/" + symb + "/history?p=" + symb
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    table = soup.find_all('table', attrs={"class": "W(100%) M(0)"})
    table_data = table[0].tbody.find_all('tr')
    price = float(table_data[0].find_all('td')[4].string.replace(',', ''))
    return price


def get_net_income(symb):
    url = "https://finance.yahoo.com/quote/" + symb + "/financials?p=" + symb
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    #table = soup.find_all('div', attrs={"class": "D(tbc)"})  # 55
    table = soup.find_all('div', attrs={"title": "Net Income Common Stockholders"})
    result = []
    d1 = table[0].next_sibling.next
    #result.append(int(d1.find('span').string.replace(',', '')))  # skip TTM data
    for i in range(0,4):
        d1=d1.next_sibling
        result.append(int(d1.find('span').string.replace(',', '')))
    return result


def get_common_stock_equity(symb):
    url = "https://finance.yahoo.com/quote/" + symb + "/balance-sheet?p=" + symb
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    table = soup.find_all('div', attrs={"title": "Common Stock Equity"})
    result = []
    d1 = table[0].next_sibling.next
    result.append(int(d1.find('span').string.replace(',', '')))
    for i in range(0,3):
        d1=d1.next_sibling
        result.append(int(d1.find('span').string.replace(',', '')))
    return result


def get_share_issued(symb):
    url = "https://finance.yahoo.com/quote/" + symb + "/balance-sheet?p=" + symb
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    table = soup.find_all('div', attrs={"title": "Share Issued"})
    result = []
    d1 = table[0].next_sibling.next
    result.append(int(d1.find('span').string.replace(',', '')))
    for i in range(0,3):
        d1=d1.next_sibling
        result.append(int(d1.find('span').string.replace(',', '')))
    return result


def get_statistics(symb):
    url = "https://finance.yahoo.com/quote/" + symb + "/key-statistics?p=" + symb
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    table = soup.find_all('table', attrs={"class": "W(100%) Bdcl(c)"})
    table_data = table[3].tbody.find_all('tr')
    result = {}
    result['fiscal_year_end'] = table_data[0].find_all('td')[1].string
    print('Fiscal Year Ends :' + result['fiscal_year_end'])
    result['most_recent_quarter'] = table_data[1].find_all('td')[1].string
    table_data = table[5].tbody.find_all('tr')
    result['roe'] = float(table_data[1].find_all('td')[1].string.replace('%', ''))
    table_data = table[7].tbody.find_all('tr')
    #a=table_data[5]
    #b=a.find_all('td')[1]
    result['bps'] = float(table_data[5].find_all('td')[1].string.replace(',', ''))
    return result


def get_avg_roe_bps(symb, year=4):
    net_income = get_net_income(symb)
    equity = get_common_stock_equity(symb)
    share = get_share_issued(symb)
    roe_sum = 0
    bps_sum = 0
    for i in range(0,4):
        roe_sum = roe_sum + net_income[i]/equity[i]
        bps_sum = bps_sum + equity[i]/share[i]
    result = {}
    result['roe'] = roe_sum/year * 100
    result['bps'] = bps_sum/year
    return result


names = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK_B', 'O', 'VZ', 'SPG']
symbs = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK-B', 'O', 'VZ', 'SPG']
Item = ['Price', 'Date', 'BPS', 'ROE', 'Value after 10yr', 'Earning Rate', '', '15% Buy Price', '10% Buy Price', '5% Buy Price']
mult15 = math.pow(1.15, 10)
mult10 = math.pow(1.1, 10)
mult5 = math.pow(1.05, 10)

gc = pygsheets.authorize(service_file='Yahoo-3018bb168e80.json')
sh = gc.open('ROE')
for symb, name in zip(symbs, names):
    print(name)
    try:
        worksheet = sh.worksheet('title', name)
    except:
        worksheet = sh.add_worksheet(name, rows=250, cols=20)

    Data1 = []
    p = get_price(symb)
    Data1.append(p)
    statistics = get_statistics(symb)
    Data1.append(statistics['most_recent_quarter'])
    bps = statistics['bps']
    Data1.append(bps)
    roe = statistics['roe']
    Data1.append(roe)
    v10 = bps * math.pow(1+roe/100, 10)
    Data1.append(v10)
    er = math.pow(v10/p, 0.1) - 1
    Data1.append(er*100)
    Data1. append('')
    Data1. append(v10/mult15)
    Data1. append(v10/mult10)
    Data1. append(v10/mult5)

    Data2 = []
    Data2.append('')
    Data2.append(statistics['fiscal_year_end'])
    avg = get_avg_roe_bps(symb)
    bps = avg['bps']
    Data2.append(bps)
    roe = avg['roe']
    Data2.append(roe)
    v10 = bps * math.pow(1+roe/100, 10)
    Data2.append(v10)
    er = math.pow(v10/p, 0.1) - 1
    Data2.append(er*100)
    Data2. append('')
    Data2. append(v10/mult15)
    Data2. append(v10/mult10)
    Data2. append(v10/mult5)

    worksheet.update_col(1, Item)
    worksheet.update_col(2, Data1)
    worksheet.update_col(3, Data2)

