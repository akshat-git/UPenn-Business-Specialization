'''
Different functions
'''

import pandas as pd
from ids import *
import datetime

def dates_bwn_twodates(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + datetime.timedelta(n)
def list_dates(days_length):
    sdate = datetime.date(2011,1,1)
    edate = datetime.date(2021,1,1)
    datelist = list(reversed([str(d) for d in dates_bwn_twodates(sdate,edate)]))[:days_length]
    dates = {'Date':datelist}
    return dates
def createstock(symbol, sheet, sheet2, sheet3, columnIndex, file_id, service,sheet2_id,sheet3_id,days):
    columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    column = columns[columnIndex]
    url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=1293840000&period2=1609459200&interval=1d&events=history&includeAdjustedClose=true"
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
def conditional(sheet_id, percentile, colormin, colormid, colormax, startR, endR, startC, endC, sheet_service):
    request_body_cond = {
        'requests':[
            {
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges':[
                            {
                                "sheetId": sheet04_id,
                                "startRowIndex": startR,
                                "endRowIndex": endR,
                                "startColumnIndex": startC,
                                "endColumnIndex": endC
                            }
                        ],
                        'gradientRule':{
                            "minpoint": {
                                "color": {
                                    "red": colormin[0],
                                    "green": colormin[1],
                                    "blue": colormin[2],
                                    "alpha": colormin[3]
                                },
                                "type":"MIN" 
                            },
                            "midpoint": {
                                "color": {
                                    "red": colormid[0],
                                    "green": colormid[1],
                                    "blue": colormid[2],
                                    "alpha": colormid[3]
                                },
                                "type": "PERCENT" ,
                                "value": str(percentile)
                            },
                            "maxpoint": {
                                "color": {
                                    "red": colormax[0],
                                    "green": colormax[1],
                                    "blue": colormax[2],  
                                    "alpha": colormax[3]
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
    response_date = sheet_service.spreadsheets().batchUpdate(
            spreadsheetId = file_id,
            body = request_body_cond
    ).execute()
def sheet04(sheet02_name, sheet03_name, sheet04_name, sheet04_id, file_id,symbollist,sheet_service,days):
    names_graph = {
        '':[symbollist['symbol01'],'100%','95%','90%','85%','80%','75%','70%','65%','60%','55%','50%','45%','40%','35%','30%','25%','20%','15%','10%','5%','0%']
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
        '':[symbollist['symbol02'],'=1-B3','=1-B4','=1-B5','=1-B6','=1-B7','=1-B8','=1-B9','=1-B10','=1-B11','=1-B12','=1-B13','=1-B14','=1-B15','=1-B16','=1-B17','=1-B18','=1-B19','=1-B20','=1-B21','=1-B22','=1-B23']
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