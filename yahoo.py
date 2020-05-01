import pandas as pd
import pandas_datareader as pdr
from datetime import datetime, timedelta
# from matplotlib import pyplot as plt
from pandas import ExcelWriter
import pygsheets
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
    df['52wkHighChg'] = (df['Close']-df['52wkHigh'])/df['52wkHigh']*100


def concat_sorted_df(df):
    dfs = df.copy()
    idx = list(dfs.index)
    for i in idx:
        s = dfs.loc[i]
        s = s.sort_values(ascending=False)
        idy = list(s.index)
        k = 0
        for j in idy:
            k = k + 1
            dfs.loc[i, j] = k
    return pd.concat([df, dfs], axis=1)


def add_chart(workbook, worksheet, key):
    chart = workbook.add_chart({'type': 'line'})
    # bold = workbook.add_format({'bold': 1})
    for i in range(len(names)):
        chart.add_series({
            'name': [key, 0, i + 1],
            'categories': [key, 1, 0, BACK, 0],
            'values': [key, 1, i + 1, BACK, i + 1],
        })

    # Add a chart title and some axis labels.
    chart.set_title({'name': key + ' History'})
    chart.set_x_axis({'name': 'Date'})
    chart.set_y_axis({'name': key + '(%)'})

    # Set an Excel chart style. Colors with white outline and shadow.
    chart.set_style(10)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('C3', chart, {'x_offset': 25, 'y_offset': 10})


# start = datetime(2019,1,1)
end = datetime.now()
start = end - timedelta(weeks=WKS)
df = pd.DataFrame()
df1 = pd.DataFrame()
for name, symb in zip(names, symbs):
    globals()[name] = pdr.get_data_yahoo(symb, start, end)
    calc_52wkratio(globals()[name])
    df[name] = globals()[name]['52wkRatio']
    df1[name] = globals()[name]['52wkHighChg']

df = concat_sorted_df(df[-1*BACK:])
df1 = concat_sorted_df(df1[-1*BACK:])

writer = ExcelWriter('corona.xlsx', datetime_format='mmm dd')
df.to_excel(writer, '52wkRatio')
df1.to_excel(writer, '52wkHighChg')

workbook = writer.book
add_chart(workbook,  writer.sheets['52wkRatio'], '52wkRatio')
add_chart(workbook,  writer.sheets['52wkHighChg'], '52wkHighChg')
workbook.close()
#writer.save()

df['DATE'] = df.index.date
df.set_index('DATE', drop=True, inplace=True)
df1['DATE'] = df1.index.date
df1.set_index('DATE', drop=True, inplace=True)
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name('Yahoo-3018bb168e80.json', scope)
client = gspread.authorize(creds)
sh = client.open('Corona')
sh.del_worksheet(sh.get_worksheet(1))
sh.del_worksheet(sh.get_worksheet(1))
worksheet = sh.add_worksheet(title='52wkRatio', rows=str(2*BACK), cols=str(3*len(name)))
set_with_dataframe(worksheet, df, include_index=True)
worksheet = sh.add_worksheet(title='52wkHighChg', rows=str(2*BACK), cols=str(3*len(name)))
set_with_dataframe(worksheet, df1, include_index=True)

gc = pygsheets.authorize(service_file='Yahoo-3018bb168e80.json')
data = []
for i in range(len(names)):
    data.append(((1,i+2), (BACK,i+2)))
sh = gc.open('Corona')
worksheet = sh[1]
chart = pygsheets.Chart(worksheet, ((1,1), (BACK,1)), ranges=data, chart_type=pygsheets.ChartType.LINE, title='52wkRatio History',anchor_cell='B2')
worksheet = sh[2]
chart = pygsheets.Chart(worksheet, ((1,1), (BACK,1)), ranges=data, chart_type=pygsheets.ChartType.LINE, title='52wkHighChg History',anchor_cell='B2')

