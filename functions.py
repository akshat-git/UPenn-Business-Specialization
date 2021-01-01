import pandas as pd
from ids import *
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
            },
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

def sheetclear(service):
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
    service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_clear
    ).execute()
