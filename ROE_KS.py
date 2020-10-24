import math
import pygsheets
import urllib.request
from bs4 import BeautifulSoup


def get_latest_roe(table):
    table_data = table[4].tbody.find_all('tr')
    roe_data = table_data[17].find_all('td')
    return float(roe_data[6].string)


def get_avg_roe(table, year=5):
    table_data = table[5].tbody.find_all('tr')
    roe_data = table_data[17].find_all('td')
    sum = 0
    for i in range(0,5):
        sum = sum + float(roe_data[i].string)
    return sum/year


def get_latest_bps(table):
    table_data = table[4].tbody.find_all('tr')
    bps_data = table_data[19].find_all('td')
    return int(bps_data[6].string.replace(',', ''))


def get_avg_bps(table, year=5):
    table_data = table[5].tbody.find_all('tr')
    roe_data = table_data[19].find_all('td')
    sum = 0
    for i in range(0,5):
        sum = sum + int(roe_data[i].string.replace(',', ''))
    return sum/year


def get_quater_date(table):
    table_data = table[4].thead.find_all('th')
    return table_data[9].string


def get_anual_date(table, year=5):
    table_data = table[4].thead.find_all('th')
    return table_data[5].string


def get_table(symb):
    url = "http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A" + symb + "&cID=&MenuYn=Y&ReportGB=&NewMenuID=101&stkGb=701"
    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    price = soup.find('span', {'id': "svdMainChartTxt11"})
    price = int(price.string.replace(',', ''))
    table = soup.find_all('table', attrs={"class": "us_table_ty1 h_fix zigbg_no"})
    return table, price


names = ['SEC', 'CJ', 'KOREIT', 'LGChem', '쌍용양회', 'CACAO', 'NAVER', 'SAMBA', 'KAIT']
symbs = ['005930', '097950', '034830', '051910', '003410', '035720', '035420', '207940', '123890']


Item = ['Price', 'Date', 'BPS', 'ROE(%)', 'Value after 10yr', '연 수익률(%)', '', '15% Buy Price', '10% Buy Price', '5% Buy Price']
mult15 = math.pow(1.15, 10)
mult10 = math.pow(1.1, 10)
mult5 = math.pow(1.05, 10)

gc = pygsheets.authorize(service_file='Yahoo-3018bb168e80.json')
sh = gc.open('ROE_KS')
for symb, name in zip(symbs, names):
    print(name)
    try:
        worksheet = sh.worksheet('title', name)
    except:
        worksheet = sh.add_worksheet(name, rows=250, cols=20)

    Data1 = []
    table, p = get_table(symb)
    Data1.append(p)
    Data1.append(get_quater_date(table))
    bps = get_latest_bps(table)
    Data1.append(bps)
    roe = get_latest_roe(table)
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
    Data2.append(get_anual_date(table))
    bps = get_avg_bps(table)
    Data2.append(bps)
    roe = get_avg_roe(table)
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

