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
# File and tickers
sheet01_name = 'Historic Prices'
sheet02_name = 'Returns'
sheet03_name = 'Risk-Returns'
sheet04_name = 'Final Report'

symbolin1 = input('Ticker Input: ')
symbolin2 = input('Ticker Input: ')

symbols = {
    "symbol01" : symbolin1, 
    "symbol02" : symbolin2
}
days = int(round(int(input("Working days: "))/5,0))
skiprows = 0

# =============================================================================
# Use tickers and put them in sheet
def createstock(symbol, sheet, sheet2, sheet3, columnIndex, file_id, service,sheet2_id,sheet3_id,days):
    columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    column = columns[columnIndex]
    url = "https://query1.finance.yahoo.com/v7/finance/download/"+ symbol +"?period1=1451520000&period2=1609372800&interval=1wk&events=history&includeAdjustedClose=true"
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



    days_returns = "=('" + sheet01_name + "'!" + column + "3-'" + sheet01_name + "'!" + column + "4)/'" + sheet01_name + "'!" + column + "4"
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
    ticker = {
        '':[symbol]
    }
    sheet_input_df = pd.DataFrame.from_dict(ticker)
    response_date = service.spreadsheets().values().update(
        spreadsheetId = file_id,
        valueInputOption = 'USER_ENTERED',
        range = sheet2+'!'+column+'2',
        body = dict(
            majorDimension = 'ROWS',
            values = sheet_input_df.T.reset_index().T.values.tolist())
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
                                        'formulaValue': formulas[symbol][2]
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
    ticker = {
        '':[symbol]
    }
    sheet_input_df = pd.DataFrame.from_dict(ticker)
    response_date = service.spreadsheets().values().update(
        spreadsheetId = file_id,
        valueInputOption = 'USER_ENTERED',
        range = sheet2+'!'+column+'2',
        body = dict(
            majorDimension = 'ROWS',
            values = sheet_input_df.T.reset_index().T.values.tolist())
    ).execute()


# =============================================================================
# Calling the function to populate the sheets with ALL values
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
request_body_cond = {
    'requests':[
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges':[
                        {
                            "sheetId": sheet04_id,
                            "startRowIndex": 2,
                            "endRowIndex": 23,
                            "startColumnIndex": 5,
                            "endColumnIndex": 6
                        }
                    ],
                    'gradientRule':{
                        "minpoint": {
                            "color": {
                                "red": 0.3,
                                "green": 0,
                                "blue": 0,
                                "alpha": 0.3
                            },
                            "type":"MIN" 
                        },
                        "midpoint": {
                            "color": {
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                                "alpha": 1
                            },
                            "type": "PERCENT" ,
                            "value": "75"
                        },
                        "maxpoint": {
                            "color": {
                                "red": 0,
                                "green": 1,
                                "blue": 0.2,  
                                "alpha": 1
                            },
                            "type": "MAX" 
                        }
                    }
                },
                'index': 0
            }
        }
    ]
}


def formatCells(startR, endR, startC, endC, sheetid, colors):
    request_body_format_cells = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                            'sheetId': sheetid,
                            'startRowIndex': startR,
                            'endRowIndex': endR,
                            'startColumnIndex': startC,
                            'endColumnIndex': endC
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor':{
                                "red": colors[0],
                                "green": colors[1],
                                "blue": colors[2],  
                                "alpha": 1
                            }
                        }

                    },
                    'fields':'*'
                }
            }
        ]
    }
    response_date = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_format_cells
    ).execute()
lightblue = [143,121,0]

# formatCells(startR, endR, startC, endC, sheetid, colors)
# formatCells(2, 5, 1, 2, sheet03_id, lightblue)





names_graph = {
    '':[symbols['symbol01'],'100%','95%','90%','85%','80%','75%','70%','65%','60%','55%','50%','45%','40%','35%','30%','25%','20%','15%','10%','5%','0%']
}
sheet_input_df = pd.DataFrame.from_dict(names_graph)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!B1',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()
names_graph = {
    '':[symbols['symbol02'],'=1-B3','=1-B4','=1-B5','=1-B6','=1-B7','=1-B8','=1-B9','=1-B10','=1-B11','=1-B12','=1-B13','=1-B14','=1-B15','=1-B16','=1-B17','=1-B18','=1-B19','=1-B20','=1-B21','=1-B22','=1-B23']
}  
sheet_input_df = pd.DataFrame.from_dict(names_graph)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!C1',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()

stdev01 ="'"+sheet03_name+"'!$C$4" 
stdev02 ="'"+sheet03_name+"'!$D$4"
sheet02_name_quote = "'"+sheet02_name+"'"
names_graph = {
    '':['St. Dev.',
    "=SQRT((B3*" + stdev01 + ")^2+(C3*" + stdev02 + ")^2+(2*COVARIANCE.P(" + sheet02_name_quote + "!$B$4:" + sheet02_name_quote + "!$B$"+ str(days+2) + "," + sheet02_name_quote + "!$C$4:" + sheet02_name_quote + "!$C$"+ str(days+2) + ")*B3*C3))"]
}   
sheet_input_df = pd.DataFrame.from_dict(names_graph)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!D1',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()
request_body_stdev_col = {
    'requests':[
        {
            'repeatCell':{
                'range':{
                    "sheetId": sheet04_id,
                    "startRowIndex": 2,
                    "endRowIndex": 23,
                    "startColumnIndex": 3,
                    "endColumnIndex": 4
                },
                'cell':{
                    "userEnteredValue": {
                        "formulaValue": names_graph[''][1]
                    },
                    "userEnteredFormat": {
                        'numberFormat':{
                            'type':'PERCENT',
                            'pattern':'0.##%'
                        },
                        'borders':{
                            "top":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "bottom":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "right":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "left":{
                                "style": 'SOLID_MEDIUM'
                            }
                        } 
                    }
                },
                'fields':'*'
            }
        }

    ]
}
response_date = sheet_service.spreadsheets().batchUpdate(
    spreadsheetId = file_id,
    body = request_body_stdev_col
).execute()
sheet03_name_quote = "'"+sheet03_name+"'!"
names_graph = { 
    '':['Returns',"=sumproduct(B3:C3,"+sheet03_name_quote+"$C$3:"+sheet03_name_quote+"$D$3)"]
}   
sheet_input_df = pd.DataFrame.from_dict(names_graph)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!E1',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()
request_body_return_col = {
    'requests':[
        {
            'repeatCell':{
                'range':{
                    "sheetId": sheet04_id,
                    "startRowIndex": 2,
                    "endRowIndex": 23,
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                },
                'cell':{
                    "userEnteredValue": {
                        "formulaValue": names_graph[''][1]
                    },
                    "userEnteredFormat": {
                        'numberFormat':{
                            'type':'PERCENT',
                            'pattern':'0.##%'
                        },
                        'borders':{
                            "top":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "bottom":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "right":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "left":{
                                "style": 'SOLID_MEDIUM'
                            }
                        } 
                    }
                },
                'fields':'*'
            }
        }
    ]
}
response_date = sheet_service.spreadsheets().batchUpdate(
    spreadsheetId = file_id,
    body = request_body_return_col
).execute()
names_graph = {
    '':['Sharpe','=E3/D3']
}
sheet_input_df = pd.DataFrame.from_dict(names_graph)
response_date = sheet_service.spreadsheets().values().update(
    spreadsheetId = file_id,
    valueInputOption = 'USER_ENTERED',
    range = sheet04_name+'!F1',
    body = dict(
        majorDimension = 'ROWS',
        values = sheet_input_df.T.reset_index().T.values.tolist())
).execute()
request_body_sharpe_col = {
    'requests':[
        {
            'repeatCell':{
                'range':{
                    "sheetId": sheet04_id,
                    "startRowIndex": 2,
                    "endRowIndex": 23,
                    "startColumnIndex": 5,
                    "endColumnIndex": 6
                },
                'cell':{
                    "userEnteredValue": {
                        "formulaValue": names_graph[''][1]
                    },
                    "userEnteredFormat": {
                        'numberFormat':{
                            'type':'NUMBER',
                            'pattern':'0.###'
                        },
                        'borders':{
                            "top":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "bottom":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "right":{
                                "style": 'SOLID_MEDIUM'
                            },
                            "left":{
                                "style": 'SOLID_MEDIUM'
                            }
                        }  
                    }
                },
                'fields':'*'
            }
        }
    ]
}
response_date = sheet_service.spreadsheets().batchUpdate(
    spreadsheetId = file_id,
    body = request_body_sharpe_col
).execute()

response_date = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_cond
).execute()





# createstock(symbol, sheet, columnIndex, file_id, service,sheet2_id,sheet3_id,days)
j = 1
for i in symbols.values():
    createstock(i, sheet01_name, sheet02_name, sheet03_name, j, file_id, sheet_service, sheet02_id, sheet03_id, days)
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


draw_chart = 0 #input("Draw chart? :")
if draw_chart == 'y':
    chart_prop = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body = request_body
    ).execute()
# chart_id = chart_prop['replies'][0]['addChart']['chart']['chartId']
# =============================================================================

request_body_clear = {
    'requests':[
        {
            'updateCells':{
                'fields':'*',
                'range':{
                    'sheetId': sheet01_id,
                }
            }
        },
        {
            'updateCells':{
                'fields':'*',
                'range':{
                    'sheetId': sheet02_id,
                    
                }
            }
        },
        {
            'updateCells':{
                'fields':'*',
                'range':{
                    'sheetId': sheet03_id,
                }
            }
        },
        {
            'updateCells':{
                'fields':'*',
                'range':{
                    'sheetId': sheet04_id,
                }
            }
        },
        {
            'deleteConditionalFormatRule':{
                'index': 0,
                'sheetId': sheet04_id
            }
        }
    ]
}

# =============================================================================
# End of program. Should I clear the sheet?
clear_sheet = input("Clear Sheet? ")
if clear_sheet == 'y':
    sheet_service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_clear
    ).execute()


# =============================================================================