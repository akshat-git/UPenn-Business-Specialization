from Google import Create_Service
import pandas as pd
# from ids import file_id
from ids import target_folder_id
from ids import file_id
from ids import sheet01_id
from ids import sheet02_id
from ids import sheet03_id 
from ids import sheet04_id
from ids import CLIENT_SECRET_FILE_DRIVE, CLIENT_SECRET_FILE_SHEET
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

# =============================================================================
# Create a new folder for keeping all the files
folder_metadata = {
    'name': 'my_folder_for_sheets',
    'mimeType': 'application/vnd.google-apps.folder'
}
# folder_prop = drive_service.files().create(body=folder_metadata).execute()
# target_folder_id = folder_prop['id']
# =============================================================================

# =============================================================================
# Create a new file with multiple sheets
sheet01_name = 'Historic Prices'
sheet02_name = 'Returns'
sheet03_name = 'Risk-Returns'
sheet04_name = 'Final Report'

file_metadata = {
    'properties': {
        'title': 'Stock Portfolio Analysis',
        'locale': 'en_US',
        'timeZone': 'America/Los Angeles',
        'autoRecalc': 'ON_CHANGE'
    },
    'sheets': [
        {
            'properties': {
                'title': sheet01_name
            }
        },
        {
            'properties': {
                'title': sheet02_name
            }
        },
        {
            'properties': {
                'title': sheet03_name
            }
        },
        {
            'properties': {
                'title': sheet04_name
            }
        }
    ]
}
# file_prop = sheet_service.spreadsheets().create(body=file_metadata).execute()
# file_id = file_prop['spreadsheetId']
# =============================================================================

# =============================================================================
#  Move the files created into this folder
'''
drive_service.files().update(
    fileId=file_id,
    addParents=target_folder_id
).execute()
'''
# =============================================================================

# =============================================================================
# Accept Tickers and put them in sheet
symbols = {
    "symbol01" : 'MSFT', 
    "symbol02" : 'WFC' ,
    "symbol03" : 'AAPL' ,
    "symbol04" : 'XOM' ,
    "symbol05" : 'COP' ,
    "symbol06" : 'BIDU' ,
    "symbol07" : 'DIS' ,
    "symbol08" : 'GOOG' ,
    "symbol09" : 'TSLA' ,
    "symbol10" : 'TTM' 
}
days = int(input("Working days: "))
skiprows = 0

def createstock(symbol, sheet, sheet3, columnIndex, file_id, service,sheet2_id,sheet3_id,days):
    columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    column = columns[columnIndex]
    url = "hist_price\\"+symbol+".csv"
    stock_price = pd.read_csv(url, skiprows = 0)
    new_set = stock_price['Close']
    new_set = dict(new_set.iloc[::-1])
    stock = {
        'Close':[]
    }
    for i in new_set.values():
        stock['Close'].append(i)
    sheetdf = pd.DataFrame.from_dict(stock)
    sheetdf.dropna(axis=0, inplace = True)
    valueslist = dict(sheetdf)['Close']
    for i in range(len(valueslist)): 
        valueslist[i] = round(valueslist[i],1)
    sheet_input = {
        symbol:valueslist[:days]
    }
    sheet_input_df = pd.DataFrame.from_dict(sheet_input)
    response_date = service.spreadsheets().values().update(
        spreadsheetId = file_id,
        valueInputOption = 'USER_ENTERED',
        range = sheet+'!'+column+'2',
        body = dict(
            majorDimension = 'ROWS',
            values = sheet_input_df.T.reset_index().T.values.tolist())
    ).execute()
    days_returns = "=('" + sheet01_name + "'!" + column + "4-'" + sheet01_name + "'!" + column + "3)/'" + sheet01_name + "'!" + column + "3"
    request_body_return = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                            'sheetId': sheet2_id,
                            'startRowIndex': 3,
                            'endRowIndex': days+2,
                            'startColumnIndex': columnIndex,
                            'endColumnIndex': columnIndex+1
                    },
                    'cell': {
                        'userEnteredValue': {
                            'formulaValue': days_returns
                        },
                        'userEnteredFormat': {
                            'numberFormat':{
                                'type': 'PERCENT',
                                'pattern': '0.##%'
                            }
                        }
                    },
                    'fields':'*'
                }
            }
        ]
    }
    response_date = service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_return
    ).execute()

    

    formulas = {
        symbol : [
            "=average('" + sheet02_name + "'!" + column + "4:"+ column + str(days+2) + ")" ,
            "=stdev.p('" + sheet02_name + "'!" + column + "4:"+ column + str(days+2) + ")" ,
            "='"+sheet03_name+"'!" + columns[columnIndex+1] + "3/'"+sheet03_name+"'!" + columns[columnIndex+1] + "4"
        ]
    }

    sheet_input_df = pd.DataFrame.from_dict(formulas)
    response_date = service.spreadsheets().values().update(
        spreadsheetId = file_id,
        valueInputOption = 'USER_ENTERED',
        range = sheet3+'!'+columns[columnIndex+1]+'2',
        body = dict(
            majorDimension = 'ROWS',
            values = sheet_input_df.T.reset_index().T.values.tolist())
    ).execute()
    request_body_formats = {
        'requests' : [
            {
                'updateCells':{
                    "rows": [
                        {
                            'values': [
                                {
                                    'userEnteredValue': {
                                        'formulaValue': formulas[symbol][1]
                                    },
                                    'userEnteredFormat': {
                                        'numberFormat':{
                                            'type': 'PERCENT',
                                            'pattern': '0.##%'
                                        }
                                    }
                                }
                            ]   
                        }
                    ],
                    "fields": '*',
                    "start": {  
                        "sheetId": sheet3_id,
                        "rowIndex": 3,
                        "columnIndex": columnIndex+1
                    }
                }
            }
        ]
    }
    response_date = service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_formats
    ).execute()

    request_body_formats = {
        'requests' : [
            {
                'updateCells':{
                    "rows": [
                        {
                            'values': [
                                {
                                    'userEnteredValue': {
                                        'formulaValue': formulas[symbol][1]
                                    },
                                    'userEnteredFormat': {
                                        'numberFormat':{
                                            'type': 'NUMBER',
                                            'pattern': '0.####'
                                        }
                                    }
                                }
                            ]   
                        }
                    ],
                    "fields": '*',
                    "start": {  
                        "sheetId": sheet3_id,
                        "rowIndex": 4,
                        "columnIndex": columnIndex+1
                    }
                }
            }
        ]
    }
    response_date = service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_formats
    ).execute()

# (symbol, sheet, columnIndex, file_id, service,sheet2_id,sheet3_id,days)

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

j = 1
for i in symbols.values():
    createstock(i, sheet01_name, sheet03_name, j, file_id, sheet_service, sheet02_id, sheet03_id, days)
    j += 1

# =============================================================================

# =============================================================================
# Plot chart
sheet_id = sheet01_id
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
                                    'title': 'Time'
                                },
                                # y-axis
                                {
                                    'position': "LEFT_AXIS",
                                    'title': 'Stock Price'
                                }
                            ],
                            # Chart data
                            'series': [
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 1, # column B
                                                    'endColumnIndex': 2
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 2, # column C
                                                    'endColumnIndex': 3
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 3, # column D
                                                    'endColumnIndex': 4
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 3, # column E
                                                    'endColumnIndex': 4
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 4, # column F
                                                    'endColumnIndex': 5
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 5, # column G
                                                    'endColumnIndex': 6
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 6, # column H
                                                    'endColumnIndex': 7
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 7, # column I
                                                    'endColumnIndex': 8
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 8, # column J
                                                    'endColumnIndex': 9
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                },
                                {
                                    'series': {
                                        'sourceRange': {
                                            'sources': [
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 2, # Row # 1
                                                    'endRowIndex': days+2, # Row # 10
                                                    'startColumnIndex': 9, # column K
                                                    'endColumnIndex': 10
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                }
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
                            'offsetXPixels': 1000,
                            'offsetYPixels': 0,
                            'widthPixels': 600,
                            'heightPixels': 400
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

# chart_id = chart_prop['replies'][0]['addChart']['chart']['chartId']

# =============================================================================
# End of program. Should I clear the sheet?
clear_sheet = input("Clear Sheet? ")
if clear_sheet == 'y':
    sheet_service.spreadsheets().values().clear(
        spreadsheetId = file_id,
        range = sheet01_name
    ).execute()

    sheet_service.spreadsheets().values().clear(
        spreadsheetId = file_id,
        range = sheet02_name
    ).execute()

    sheet_service.spreadsheets().values().clear(
        spreadsheetId = file_id,
        range = sheet03_name
    ).execute()
# =============================================================================
