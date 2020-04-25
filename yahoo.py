import pandas as pd
import pandas_datareader as pdr
from datetime import datetime, timedelta
from matplotlib import pyplot as plt
from pandas import ExcelWriter
import pygsheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe


BACK = 30  # How many days we go back to set start day for output
WKS = 52*2


def calc_52wkratio(df):
    df['52wkHigh'] = df['High'].asfreq('D').rolling(window=52*7, min_periods=1).max()
    df['52wkLow'] = df['Low'].asfreq('D').rolling(window=52*7, min_periods=1).min()
    df['52wkRatio'] = (df['Close']-df['52wkLow'])/(df['52wkHigh']-df['52wkLow'])*100


# start = datetime(2019,1,1)
end = datetime.now()
start = end - timedelta(weeks=WKS)
DAL = pdr.get_data_yahoo('DAL', start, end)
GOOGL = pdr.get_data_yahoo('GOOGL', start, end)
AAPL = pdr.get_data_yahoo('AAPL', start, end)
T = pdr.get_data_yahoo('T', start, end)
MO = pdr.get_data_yahoo('MO', start, end)
BRK_B = pdr.get_data_yahoo('BRK-B', start, end)
XOM = pdr.get_data_yahoo('XOM', start, end)
O = pdr.get_data_yahoo('O', start, end)
SPG = pdr.get_data_yahoo('SPG', start, end)
VZ = pdr.get_data_yahoo('VZ', start, end)
EMB = pdr.get_data_yahoo('EMB', start, end)
IDV = pdr.get_data_yahoo('IDV', start, end)
GSPC = pdr.get_data_yahoo('^GSPC', start, end)

calc_52wkratio(DAL)
calc_52wkratio(GOOGL)
calc_52wkratio(AAPL)
calc_52wkratio(T)
calc_52wkratio(MO)
calc_52wkratio(BRK_B)
calc_52wkratio(O)
calc_52wkratio(SPG)
calc_52wkratio(VZ)
calc_52wkratio(EMB)
calc_52wkratio(IDV)
calc_52wkratio(GSPC)

print(GSPC.tail())
#GSPC['52wkRatio'][-1*BACK:].plot()
#plt.show()

df = pd.DataFrame({'DAL':DAL['52wkRatio'],
                   'GOOGL':GOOGL['52wkRatio'],
                   'AAPL':AAPL['52wkRatio'],
                   'T':T['52wkRatio'],
                   'MO':MO['52wkRatio'],
                   'BRK_B':BRK_B['52wkRatio'],
                   'O':O['52wkRatio'],
                   'SPG':SPG['52wkRatio'],
                   'VZ':VZ['52wkRatio'],
                   'EMB':EMB['52wkRatio'],
                   'IDV':IDV['52wkRatio'],
                   'GSPC':GSPC['52wkRatio']})

# ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK_B', 'O', 'SPG', 'VZ', 'EMB', 'IDV', 'GSPC']
print(df.tail())
writer = ExcelWriter('corona.xlsx')
df[-1*BACK:].to_excel(writer, '52wkRatio')
writer.save()

"""
gc = pygsheets.authorize(service_file='Yahoo-3018bb168e80.json')
sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/141_KPbv0ZvcC2-RjwUOMxVAxgD0QTFAboT7E8ty-UgM/edit?usp=sharing')
wks = sh[0]
wks.set_dataframe(df[-1*BACK:])
"""
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']  #'https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name('Yahoo-3018bb168e80.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Corona').sheet1
#sheet.update_cell(1, 1, 'test')
set_with_dataframe(sheet, df[-1*BACK:], include_index=True)

