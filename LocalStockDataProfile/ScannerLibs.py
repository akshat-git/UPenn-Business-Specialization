import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import SharpeCalc
import CommonStocks

def Reverse(lst): 
    return [ele for ele in reversed(lst)] 

def scan_stock(ticker, days, columnIndex, nameIndex):
    columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    column = columns[columnIndex]
    url = "https://query1.finance.yahoo.com/v7/finance/download/" + ticker + "?period1=1293840000&period2=1609459200&interval=1d&events=history&includeAdjustedClose=true"
    stock_price = pd.read_csv(url, skiprows = 0)
    new_set = stock_price[nameIndex]
    new_set = dict(new_set.iloc[::-1])
    stock = {
        nameIndex:[]
    }
    for i in new_set.values():
        stock[nameIndex].append(i)
    lst = []
    dlist = stock[nameIndex]
    for i in range(0, days):
        lst.append(dlist[(days - i)])
    return lst

def date_list(vnum):
    dl = []
    for i in range(0, vnum):
        dl.append((i+1)-vnum)
    return dl


def retrieve_stock(ticker, days, showplot = True):
    plt.plot(date_list(days), scan_stock(ticker, days, 5, 'Close'))
    if showplot:
        plt.show()
    return plt

def project_stock(stock, x, y, days, showplot = True):
    x = np.array(x)
    y = np.array(y) 
    m, b = np.polyfit(x, y, 1)
    plt.plot(x, m*x + b)
    retrieve_stock(stock, days, False)
    if showplot:
        plt.show()
    plt.cla()
    plt.plot(x, m*x + b)
    return (x, (m*x+b), m, b)


def stock_returns_over_interval(stock, sd, ed, prevhist, prex = False):
    npxarr, npyarr, m, b = project_stock(stock, date_list(prevhist), scan_stock(stock, prevhist, 5, 'Close'), 180, False)
    ret = (m * ed + b) - (m * sd + b)
    return ret

def print_stock_returns_to_terminal(stock):
    print("Expected returns over 18 day interval is: $" + str(stock_returns_over_interval(stock, 0, 18, 360)))
    print("With recent data: $" + str(stock_returns_over_interval(stock, 0, 18, 36)))

def stock_sharpe_ratio(ticker, time, prevhdays):
    return SharpeCalc.sharpe(ticker, time, prevhdays)

def project_compare_stocks(tickers, bd, sd, dph = 180):
    hs_proj = None
    hs_proj_ret = 0
    for i in range(0, len(tickers)):
        lr_ret = stock_returns_over_interval(tickers[i], bd, sd, dph)
        if lr_ret > hs_proj_ret:
            hs_proj = tickers[i]
            hs_proj_ret = lr_ret
    hs_sharpe = None
    hs_sharpe_ret = 0
    for i in range(0, len(tickers)):
        lr_ret = stock_sharpe_ratio(tickers[i], (sd-bd), dph)
        if lr_ret > hs_sharpe_ret:
            hs_sharpe = tickers[i]
            hs_sharpe_ret = lr_ret
    hs_risk = None
    hs_risk_ret = (10^1000)
    for i in range(0, len(tickers)):
        lr_ret = SharpeCalc.risk(tickers[i], dph)
        if lr_ret < hs_risk_ret:
            hs_risk = tickers[i]
            hs_risk_ret = lr_ret
    return hs_proj, hs_sharpe, hs_risk

def calculate_buy_or_sell(ticker, pvh):
    npxarr, npyarr, m, b = project_stock(ticker, date_list(pvh), scan_stock(ticker, pvh, 5, 'Close'), 180, False)
    stock_hist = scan_stock(ticker, pvh, 5, 'Close')
    stock_date = date_list(pvh)
    if stock_hist[len(stock_hist)-1] > b + stock_hist[len(stock_hist)-1]/30:
        return 'sell'
    if stock_hist[len(stock_hist)-1] > b - stock_hist[len(stock_hist)-1]/60:
        return 'buy'
    else:
        return 'none'