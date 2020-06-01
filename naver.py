from  yahoo import *

names = ['SEC', 'SEC_P', 'INFRA', 'US_REIT', 'SP500', 'CJ', 'KOREIT', 'GOLD', 'USD', '10Y_KS', '10Y_US']
codes = ['005930', '005935', '329200', '182480', '219480', '097950', '034830', '132030', '132030', '148070', '305080']


def get_data_naver(code, start, end, max_page=20):
    df = pd.DataFrame()
    url = 'https://finance.naver.com/item/sise_day.nhn?code={code}'.format(code=code)
    print("Requested URL = {}".format(url))
    for page in range(1, max_page):
        page_url = '{url}&page={page}'.format(url=url, page=page)
        df = df. append(pd.read_html(page_url, header=0)[0], ignore_index=True)
    df = df.dropna()
    df = df[::-1]
    df = df[['날짜', '종가', '시가', '고가', '저가']]
    df.rename(columns={'날짜':'Date', '종가':'Close', '시가':'Open', '고가':'High', '저가':'Low'}, inplace=True)
    df['Date'] = pd.to_datetime(df['Date'])
    df.set_index(pd.DatetimeIndex(df['Date'].values), drop=True, inplace=True)
    df.drop('Date', axis=1,inplace=True)
    return df


main(names, codes, 'corona_ks.xlsx', 'Corona_ks', get_data_naver)

