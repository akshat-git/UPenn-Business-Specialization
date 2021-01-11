'''
This file contains the main code
    - uses google spreadsheet api instance
    - takes in input for ticker and days
    - puts values of prices in sheet1
    - puts returns in sheet2
    - puts stats in sheet 3
    - puts graph and data for each set of weightages in sheet4
'''

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

# importform(service, spreadsheet_id,range):
symbollist = importform(sheet_service,file_id, "'Form Responses 1'!B2:E2")['values'][0]
print(symbollist)
symbolin1 = 'FB'    #input('Ticker Input: ')
symbolin2 = 'ZM'    #input('Ticker Input: ')
symbolin3 = 'TSLA'  #input('Ticker Input: ')
symbolin4 = 'WFC'   #input('Ticker Input: ')
symbols = {
    "symbol01" : symbollist[0], 
    "symbol02" : symbollist[1],
    "symbol03" : symbollist[2], 
    "symbol04" : symbollist[3]

}
days = 100 #int(input("Working days: "))
skiprows = 0

# =============================================================================
# Names sheet 03
names = {
    '':['Returns','St. Dev.','Sharpe']
}
sheet_input_df = pd.DataFrame.from_dict(names)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet03_name+'!B2',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()





# formatCells(range, sheetid, colors)
# formatCells(range, sheet03_id, lightblue)

# conditional(sheet_id, percentile, colormin, colormid, colormax, range, service)
conditional(sheet04_id, 99.99, colors['white'], colors['lightgreen'], colors['green'],named_ranges['sharpe'], sheet_service)

date_col = list_dates(days,symbolin1)
date_df = pd.DataFrame.from_dict(date_col)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet01_name+'!A2',
    body = dict(
        majorDimension = 'ROWS',
        values = date_df.T.reset_index().T.values.tolist())
).execute()

# createstock(symbol, sheet, columnIndex, file_id, service,sheet2_id,sheet3_id,days)
j = 1
for i in symbols.values():
    createstock(i, sheet01_name, sheet02_name, sheet03_name, j, file_id, sheet_service, sheet02_id, sheet03_id, days)
    j += 1

# sheet04(sheet02_name, sheet03_name, sheet04_name, sheet04_id, file_id)
sheet04(sheet02_name, sheet03_name, sheet04_name, sheet04_id, file_id, symbols,sheet_service,days)


tangline = {
    '0':['=MAX(F3:F23)'],
    '0':['=MAX(D3:D23)']
}
'''
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


# =============================================================================
# Plot chart

draw_chart = 'y' #input("Draw chart? :")
if draw_chart == 'y':
    #chart_draw(service, sheet_id, domain, series,type)
    chart_id = chart_draw(sheet_service, sheet04_id,named_ranges['stdev'],named_ranges['returns'],'LINE')
    chart_id_bubble = chart_draw_bubble(sheet_service, sheet03_id)
# =============================================================================

# =============================================================================
# End of program. Should I clear the sheet?
clear_sheet = input("Clear Sheet? ")
if clear_sheet == 'y':
    sheetclear(sheet_service, chart_id)
    sheetclear(sheet_service, chart_id_bubble)
# =============================================================================