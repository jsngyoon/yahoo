import math
import pygsheets
# import urllib.request
from bs4 import BeautifulSoup
import requests
# from random import choice
# import xlsxwriter

# BaseURL			= "https://finance.yahoo.com/quote/ACN/financials?p=ACN"
BSURL = "https://finance.yahoo.com/quote/{0}/balance-sheet?p={0}"
ISURL = "https://finance.yahoo.com/quote/{0}/financials?p={0}"
CFURL = "https://finance.yahoo.com/quote/{0}/cash-flow?p={0}"
STURL = "https://finance.yahoo.com/quote/{0}/key-statistics?p={0}"


def getRequest(aurl):
    user_agent = 'Mozzila/5.0'
    hdr = {'User-Agent': user_agent}

    print("Requesting website : " + aurl)

    try:
        req = requests.get(aurl, headers=hdr)
    except Exception as err:
        print('Error in request. Error :')
        print(err.message)

    return req


def getSoup(aurl):
    req = getRequest(aurl)
    content = req.content

    soup = BeautifulSoup(content, 'html.parser')
    return soup


def get_financial_data(BaseURL, code, key_name):
    url = BaseURL.format(code)
    soup = getSoup(url)
    p = soup.prettify()
    # Identifying the section containing data
    section = soup.find('section', {'class': 'Mb(30px)'})
    records = section.find_all('div', {'class': 'D(tbr)'})

    for record in records:
        item = record.find('div', {'class': 'D(ib)'}).find('span').text
        if item == 'Breakdown':
            date = get_row_data(record)
        elif item == key_name:
            result = get_row_data(record)
    return date, result


def get_row_data(record):
    divs = record.find_all('div', {'class': 'Ta(c)'})
    date = []
    for div in divs:
        date.append(div.find('span').text)
    return date


def get_statiscal_data(code, key_name, BaseURL=STURL):
    url = BaseURL.format(code)
    soup = getSoup(url)

    price = soup.find('div', 'My(6px)', 'Pos(r)').find('span').text  # Current price

    ths = soup.table.find_all('th')
    current = ths[1].find('span', 'Pos(a)').text
    current = current.split(':')[1]  # Current date
    date = []
    date.append(current[1:])
    for th in ths[2:]:
        date.append(th.span.text)  # Quarterly date

    tables = soup.find_all('table')
    for table in tables:
        tbody = table.tbody
        trs = tbody.find_all('tr')
        for tr in trs:
            tds = tr.find_all('td')
            if tds[0].span.text == key_name or tds[0].span.text in key_name:
                result = []
                for td in tds[1:]:
                    result.append(td.text)
                return price, date, result
    print("No match found for " + key_name)
    return price, date, []


'''
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
        try:
            result.append(int(d1.find('span').string.replace(',', '')))
        except:
            print('get_net_income() for ' + symb, i)
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
        try:
            result.append(int(d1.find('span').string.replace(',', '')))
        except:
            print('get_common_stock_equity() for ' + symb, i)
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
        try:
            result.append(int(d1.find('span').string.replace(',', '')))
        except:
            print('get_share_issued() for ' + symb, i)
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
'''

def calc_bps_roe(share, equity, income):
    l = len(share)
    bps = [0]*(l+1)
    roe = [0]*(l+1)
    bps_sum = 0
    roe_sum = 0
    for i in range(0, l):
        share[i] = share[i].replace(',', '')
        equity[i] = equity[i].replace(',', '')
        income[i] = income[i].replace(',', '')
        share_ = int(share[i])
        equity_ = int(equity[i])
        income_ = int(income[i])
        bps[i+1] = equity_ / share_
        bps_sum = bps_sum + bps[i+1]
        roe[i+1] = income_ / equity_ * 100
        roe_sum = roe_sum + roe[i+1]
    bps[0] = bps_sum / l
    roe[0] = roe_sum / l
    return bps, roe

names = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK_B', 'O', 'VZ', 'SPG', 'XOM', 'TSLA', 'NVDA']
symbs = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK-B', 'O', 'VZ', 'SPG', 'XOM', 'TSLA', 'NVDA']
Item = ['Price', 'Date', 'BPS', 'ROE', 'Value after 10yr', 'Earning Rate(%)', '', '15% Buy Price', '10% Buy Price', '5% Buy Price']
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

    price, q_date, roe_ttm = get_statiscal_data(symb, 'Return on Equity (ttm)')
    _, _, bps_mrq = get_statiscal_data(symb, 'Book Value Per Share (mrq)')

    Data1 = []  # Current data
    Data1.append(price)
    Data1.append(q_date[0])
    Data1.append(bps_mrq[0])
    Data1.append(roe_ttm[0])
    bps_mrq[0] = bps_mrq[0].replace(',', '')
    bps_mrq = float(bps_mrq[0])
    roe_ttm = float(roe_ttm[0][:-2])
    price = price.replace(',', '')
    price = float(price)
    v10 = bps_mrq * math.pow(1+roe_ttm/100, 10)
    Data1.append(v10)
    er = math.pow(v10/price, 0.1) - 1
    Data1.append(er*100)
    Data1.append('')
    Data1.append(v10/mult15)
    Data1.append(v10/mult10)
    Data1.append(v10/mult5)

    date, share = get_financial_data(BSURL, symb, 'Share Issued')
    _, cse = get_financial_data(BSURL, symb, 'Common Stock Equity')
    _, income = get_financial_data(ISURL, symb, 'Net Income Common Stockholders')
    bps, roe = calc_bps_roe(share, cse, income)

    Data2 = []
    Data2.append('')
    Data2.append('Average')
    Data2.append(bps[0])
    Data2.append(roe[0])
    v10 = bps[0] * math.pow(1+roe[0]/100, 10)
    Data2.append(v10)
    er = math.pow(v10/price, 0.1) - 1
    Data2.append(er*100)
    Data2.append('')
    Data2.append(v10/mult15)
    Data2.append(v10/mult10)
    Data2.append(v10/mult5)

    worksheet.update_col(1, Item)
    worksheet.update_col(2, Data1)
    worksheet.update_col(3, Data2)

    l = len(bps)-1
    for i in range(0,l):
        Data3 = []
        Data3.append('')
        Data3.append(date[i])
        Data3.append(bps[i+1])
        Data3.append(roe[i+1])
        v10 = bps[i+1] * math.pow(1 + roe[i+1] / 100, 10)
        Data3.append(v10)
        er = math.pow(v10 / price, 0.1) - 1
        Data3.append(er * 100)
        Data3.append('')
        Data3.append(v10 / mult15)
        Data3.append(v10 / mult10)
        Data3.append(v10 / mult5)
        worksheet.update_col(i+4, Data3)



