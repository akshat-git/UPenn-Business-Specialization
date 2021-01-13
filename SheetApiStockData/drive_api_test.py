'''
This file contains the main code
    - uses google spreadsheet api instance
    - takes in input for ticker and days
    - puts values of prices in sheet1
    - puts returns in sheet2
    - puts stats in sheet 3
    - puts graph and data for each set of weightages in sheet4
'''

from Google import Create_Service
import pandas as pd
from ids import *
from functions import *
import datetime

# =============================================================================
# Start drive related services
API_NAME_DRIVE = 'drive'
API_VERSION_DRIVE = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

# Sheet related service parameters
API_NAME_SHEET = 'sheets'
API_VERSION_SHEET = 'v4'

drive_service = Create_Service(CLIENT_SECRET_FILE_DRIVE, API_NAME_DRIVE, API_VERSION_DRIVE, SCOPES)
sheet_service = Create_Service(CLIENT_SECRET_FILE_SHEET, API_NAME_SHEET, API_VERSION_SHEET, SCOPES)

# =============================================================================
# File and tickers

# Global variables (used across functions)
days = 100 #int(input("Working days: "))
colors = {
    'red':          [1,0,0,1],
    'lightred':     [0.5,0,0,0.1],
    'blue':         [0,0,1,1],
    'lightblue':    [0,0,0.5,0.1],
    'green':        [0,1,0,1],
    'lightgreen':   [0,0.5,0,0.1],
    'white':        [1,1,1,1],
}
named_ranges = {
    'stdev':[2,23,3,4],
    'returns':[2,23,4,5],
    'sharpe':[2,23,5,6]
}

# Mark the record as done reading
def MarkDone(sheet_service, file_id, symbollistmain, sheet): 
    # Mark as done
    done = {
        1:['']
    }
    donedf = pd.DataFrame.from_dict(done)
    response_date = sheet_service.spreadsheets().values().update(
        spreadsheetId = file_id,
        valueInputOption = 'USER_ENTERED',
        range = sheet01_name+'!H'+str(len(symbollistmain)-1),
        #"'Form Responses 1'!H"+str(len(symbollistmain)-1),
        body = dict(
            majorDimension = 'ROWS',
            values = donedf.T.reset_index().T.values.tolist())
    ).execute()

# Read historical data and process
def ReadHistDataAndProcess(sheet01name, sheet02name, sheet03name, fileid, sheetservice, sheet02id, sheet03id, days, symboldate, symbols):
    j = 1
    for i in symbols.values():
        # createstock(symbol, sheet, columnIndex, file_id, service,sheet2_id,sheet3_id,days)
        createstock(i, sheet01name, sheet02name, sheet03name, j, fileid, sheetservice, sheet02id, sheet03id, days)
        j += 1

    date_col = list_dates(days,symboldate)
    date_df = pd.DataFrame.from_dict(date_col)
    response_date = sheetservice.spreadsheets().values().update(
        spreadsheetId = fileid,
        valueInputOption = 'USER_ENTERED',
        range = sheet01name+'!A2',
        body = dict(
            majorDimension = 'ROWS',
            values = date_df.T.reset_index().T.values.tolist())
    ).execute()

# Create Return-Risk Names
def CreateRiskReturnNames(sheetservice, fileid, sheet03name):
    names = {
        '':['Returns','St. Dev.','Sharpe']
    }
    sheet_input_df = pd.DataFrame.from_dict(names)
    response_date = sheetservice.spreadsheets().values().update(
        spreadsheetId = fileid,
        valueInputOption = 'USER_ENTERED',
        range = sheet03name+'!B2',
        body = dict(
            majorDimension = 'ROWS',
            values = sheet_input_df.T.reset_index().T.values.tolist())
    ).execute()

# Create Final Return(Sharpe Ratio chart)
def CreateFinalReportSheet(sheetservice, fileid, sheet04id, colors, namedranges, days, sheet02name, sheet03name, sheet04name, symbols):
    # conditional(sheet_id, percentile, colormin, colormid, colormax, range, service)
    conditional(sheet04id, 99.99, colors['white'], colors['lightgreen'], colors['green'],namedranges['sharpe'], sheetservice)

    # sheet04(sheet02_name, sheet03_name, sheet04_name, sheet04_id, file_id)
    sheet04(sheet02name, sheet03name, sheet04name, sheet04id, fileid, symbols, sheetservice, days)


# EXPERIMENTAL CODE
'''
tangline = {
    '0':['=MAX(F3:F23)'],
    '0':['=MAX(D3:D23)']
}

sheet_input_df = pd.DataFrame.from_dict(tangline)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!G3',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()
'''

def main():
    ### Get the input (Ticker)
    # importform(service, spreadsheet_id,range):
    symbollistmain = importform(sheet_service, file_id, "'Form Responses 1'!B2:E")['values']
    symbollist = symbollistmain[len(symbollistmain)-1]
    symbols = {
        "symbol01" : symbollist[0], 
        "symbol02" : symbollist[1],
        "symbol03" : symbollist[2], 
        "symbol04" : symbollist[3]
    }
    print(symbollistmain)
    ### Mark the record has been read
    # MarkDone(sheet_service, file_id, symbollistmain, sheet01_id)

    ### Get historic data for stock and process the information
    ReadHistDataAndProcess(sheet01_name, sheet02_name, sheet03_name, file_id, sheet_service, sheet02_id, sheet03_id, days, symbollist[0],symbols)
    CreateRiskReturnNames(sheet_service, file_id, sheet03_name)
    # formatCells(range, sheetid, colors)
    # formatCells(range, sheet03_id, lightblue)

    ### Create sheet4 data + charts
    CreateFinalReportSheet(sheet_service, file_id, sheet04_id, colors, named_ranges, days, sheet02_name, sheet03_name, sheet04_name, symbols)
    
    ### Adding sheet03 bubb chart
    draw_chart = 'y' #input("Draw chart? :")
    if draw_chart == 'y':
        #chart_draw(service, sheet_id, domain, series,type)
        chart_id = chart_draw(sheet_service, sheet04_id,named_ranges['stdev'],named_ranges['returns'],'LINE')
        chart_id_bubble = chart_draw_bubble(sheet_service, sheet03_id)

    ### Clear sheets and charts
    clear_sheet = input("Clear Sheet? ")
    if clear_sheet == 'y':
        sheetclear(sheet_service, chart_id)
        if draw_chart == 'y':
            sheetclear(sheet_service, chart_id_bubble)

main()
