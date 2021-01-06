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

symbolin1 = input('Ticker Input: ')
symbolin2 = input('Ticker Input: ')
symbolin3 = input('Ticker Input: ')
symbolin4 = input('Ticker Input: ')
symbols = {
    "symbol01" : symbolin1, 
    "symbol02" : symbolin2,
    "symbol03" : symbolin3, 
    "symbol04" : symbolin4

}
days = int(input("Working days: "))
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





# formatCells(startR, endR, startC, endC, sheetid, colors)
# formatCells(2, 5, 1, 2, sheet03_id, lightblue)

# conditional(sheet_id, percentile, colormin, colormid, colormax, startR, endR, startC, endC, service)
conditional(sheet04_id, 99.99, colors['white'], colors['lightgreen'], colors['green'], 2, 23, 5, 6, sheet_service)

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

sheet_id = sheet04_id
request_body = {
    'requests': [
        {
            'addChart': {
                'chart': {
                    'spec': {
                        'title': 'Stock Performance',
                        'basicChart': {
                            'chartType': 'LINE',
                            'legendPosition': 'BOTTOM_LEGEND',
                            'axis': [
                                # x-axis
                                {
                                    'position': "BOTTOM_AXIS",
                                    'title': 'Standard Deviation'
                                },
                                # y-axis
                                {
                                    'position': "LEFT_AXIS",
                                    'title': 'Stock Returns'
                                }
                            ],
                            # Chart data
                            'domains':[
                                {
                                    'domain':{
                                        'sourceRange':{
                                            'sources':[
                                                {
                                                   'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': 23, # Row # 10
                                                    'startColumnIndex': 3, # column B
                                                    'endColumnIndex': 4 
                                                }
                                            ]
                                        }
                                    }
                                }
                            ],
                            'series': [
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': 23, # Row # 10
                                                    'startColumnIndex': 4, # column B
                                                    'endColumnIndex': 5
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS',                                    
                                }
                                # {
                                #     'series': {
                                #         'sourceRange': {
                                #             'sources': [
                                #                 {
                                #                     'sheetId': sheet_id,
                                #                     'startRowIndex': 2, # Row # 1
                                #                     'endRowIndex': 4, # Row # 10
                                #                     'startColumnIndex': 6, # column B
                                #                     'endColumnIndex': 7
                                #                 }
                                #             ]
                                #         }
                                #     },
                                #     'targetAxis': 'LEFT_AXIS',                                    
                                # }
                            ]
                        }
                    },
                    'position': {
                        'overlayPosition': {
                            'anchorCell': {
                                'sheetId': sheet_id,
                                'rowIndex': 1,
                                'columnIndex': 1
                            },
                            'offsetXPixels': 506,
                            'offsetYPixels': 21,
                            'widthPixels': 800,
                            'heightPixels': 466
                        }
                     }
                }
            }
        }
    ]
}


draw_chart = 'y' #input("Draw chart? :")
if draw_chart == 'y':
    chart_prop = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body = request_body
    ).execute()

chart_id = chart_prop['replies'][0]['addChart']['chart']['chartId']

# =============================================================================

# =============================================================================
# End of program. Should I clear the sheet?
clear_sheet = input("Clear Sheet? ")
if clear_sheet == 'y':
    sheetclear(sheet_service)

    # Routine to delete the embedded chart created in the previous step
    request_body = {
        'requests': [
            {
                'deleteEmbeddedObject': {
                    'objectId': chart_id
                }
            }
        ]
    }

    chart_prop = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body = request_body
    ).execute()
# =============================================================================