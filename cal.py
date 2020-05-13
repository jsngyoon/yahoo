from datetime import datetime, timedelta
import pandas as pd
import pandas_datareader as pdr
import pygsheets
# from matplotlib import pyplot as plt
from pandas import ExcelWriter
import calendar as cal
import xlsxwriter


def make_cal_df(year, month):
    monthrange = cal.monthrange(year, month)
    day = range(1, monthrange[1]+1)
    wkd = [0] * monthrange[1]
    hh = [11] * monthrange[1]
    mm = [0] * monthrange[1]
    for i in day:
        if cal.weekday(year, month, i) > 4:
            wkd[i-1] = 1
            hh[i-1] = 0
    df = pd.DataFrame(zip(wkd, day, hh, mm), columns=['Wkd', 'Day', 'HH', 'MM'])
    return df


def make_xlsx_cal_year(year):
    writer = ExcelWriter('calendar.xlsx', engine='xlsxwriter')
    for i in range(1, 13):
        df = make_cal_df(year, i)
        df.to_excel(writer, str(i), index=False)
    writer.save()


def make_gc_cal_year(gc, year):
    sh = gc.open('Calendar')
    for i in range(1, 13):
        df = make_cal_df(year, i)
        sh.del_worksheet(sh[1])
        worksheet = sh.add_worksheet(title=str(i), rows=40, cols=20)
        worksheet.set_dataframe(df, 'A1')


def make_gc_cal_month(gc, year, month):
    sh = gc.open('Working_hours')
    if month == 12:
        year = year + 1
        month = 0

    wks = sh[month]
    if wks.title == str(month + 1):
        df = make_cal_df(year, month + 1)
        wks.set_dataframe(df, 'A1')


Year = datetime.now().year
Month = datetime.now().month
# make_xlsx_cal_year(Year)
gc = pygsheets.authorize(service_file='Yahoo-3018bb168e80.json')
# make_gc_cal_year(gc, Year)
make_gc_cal_month(gc, Year, Month)

