import pandas as pd
import pandas_datareader as pdr
from datetime import datetime, timedelta
# from matplotlib import pyplot as plt
from pandas import ExcelWriter
# import pygsheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe

# BACK = 30  # How many days we go back to set start day for output
BACK = len(pd.bdate_range('2020-3-15', datetime.now()))
WKS = 52*2
names = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK_B', 'O', 'SPG', 'VZ', 'EMB', 'IDV', 'GSPC']
symbs = ['DAL', 'GOOGL', 'AAPL', 'T', 'MO', 'BRK-B', 'O', 'SPG', 'VZ', 'EMB', 'IDV', '^GSPC']


def calc_52wkratio(df):
    df['52wkHigh'] = df['High'].asfreq('D').rolling(window=52*7, min_periods=1).max()
    df['52wkLow'] = df['Low'].asfreq('D').rolling(window=52*7, min_periods=1).min()
    df['52wkRatio'] = (df['Close']-df['52wkLow'])/(df['52wkHigh']-df['52wkLow'])*100


# start = datetime(2019,1,1)
end = datetime.now()
start = end - timedelta(weeks=WKS)
df = pd.DataFrame()
for name, symb in zip(names, symbs):
    globals()[name] = pdr.get_data_yahoo(symb, start, end)
    calc_52wkratio(globals()[name])
    df[name] = globals()[name]['52wkRatio']

writer = ExcelWriter('corona.xlsx')
df[-1*BACK:].to_excel(writer, '52wkRatio')
writer.save()

scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name('Yahoo-3018bb168e80.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Corona').sheet1
set_with_dataframe(sheet, df[-1*BACK:], include_index=True)

