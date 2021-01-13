import datetime
import numpy as np
import pandas as pd
import urllib.request

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

def get_stock_profile(ticker, days):
    datelist = date_list(days)
    pricelist = scan_stock(ticker, days, 5, 'Close')
    return datelist, pricelist

def sharpe(ticker, time, prevhdays):
    sharpelist = []
    prevstock = 0
    dlist, plist = get_stock_profile(ticker, prevhdays)
    for i in range(0, len(plist)):
        if i == 0:
            prevstock = plist[i]
        else:
            midsharpe = (plist[i] - prevstock)/plist[i]
            sharpelist.append(midsharpe)
    avg = 0
    avgn = 0
    for i in range(0, len(sharpelist)):
        avgn += 1
        avg += sharpelist[i]
    std = risk(ticker, prevhdays)
    return (((avg / avgn)/std) * time)

def risk(ticker, prevhdays):
    dlist, plist = get_stock_profile(ticker, prevhdays)
    std = np.std(plist)
    return std