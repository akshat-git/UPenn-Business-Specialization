'''
Different functions
'''

import pandas as pd
from ids import *
import datetime


def importform(service, spreadsheet_id,range):
    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range, majorDimension='ROWS')
    response = request.execute()
    return response
def list_dates(days_length,symbol):
    url = "https://query1.finance.yahoo.com/v7/finance/download/" + symbol + "?period1=1293840000&period2=1609459200&interval=1d&events=history&includeAdjustedClose=true"
    stock_price = pd.read_csv(url, skiprows = 0)
    new_set = stock_price['Date']
    new_set = dict(new_set.iloc[::-1])
    date = {
        'Date':[]
    }
    for i in new_set.values():
        date['Date'].append(i)
    date['Date'] = date['Date'][:days_length]
    return date
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
def formatCells(range, sheetid, colors):
    request_body_format_cells = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                            'sheetId': sheetid,
                            "startRowIndex": range[0],
                            "endRowIndex": range[1],
                            "startColumnIndex": range[2],
                            "endColumnIndex": range[3]
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
def sheetclear(service,chartid):
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
            },
            {
                'deleteEmbeddedObject': {
                    'objectId': chartid
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_clear
    ).execute()
def conditional(sheet_id, percentile, colormin, colormid, colormax, range, sheet_service):
    request_body_cond = {
        'requests':[
            {
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges':[
                            {
                                "sheetId": sheet_id,
                                "startRowIndex": range[0],
                                "endRowIndex": range[1],
                                "startColumnIndex": range[2],
                                "endColumnIndex": range[3]
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
def chart_draw(service, sheet_id, domain, series,type):
    request_body = {
        'requests': [
            {
                'addChart': {
                    'chart': {
                        'spec': {
                            'title': 'Stock Performance',
                            'basicChart': {
                                'chartType': type,
                                'legendPosition': 'BOTTOM_LEGEND',
                                'axis': [
                                    # x-axis
                                    {
                                        'position': "BOTTOM_AXIS",
                                        'title': 'Standard Deviation',
                                        'viewWindowOptions': {
                                            'viewWindowMin': 0
                                        }
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
                                                        'startRowIndex': domain[0], # Row # 1
                                                        'endRowIndex': domain[1], # Row # 10
                                                        'startColumnIndex': domain[2], # column B
                                                        'endColumnIndex': domain[3]
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
                                                        'startRowIndex': series[0], # Row # 1
                                                        'endRowIndex': series[1], # Row # 10
                                                        'startColumnIndex': series[2], # column B
                                                        'endColumnIndex': series[3]
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
    chart_prop = service.spreadsheets().batchUpdate(
            spreadsheetId = file_id,
            body = request_body
    ).execute()
    chartidnew = chart_prop['replies'][0]['addChart']['chart']['chartId']
    return chartidnew
def chart_draw_bubble(service, sheet_id):
    request_body_bubble = {
        'requests': [
            {
                'addChart': {
                    'chart': {
                        'spec': {
                            'title': 'Reward versus Risk',
                            'bubbleChart': {
                                'legendPosition': 'BOTTOM_LEGEND',
                                # Chart data
                                'domain': {
                                    # X axis values
                                    'sourceRange':{
                                        'sources':[
                                            {
                                                'sheetId': sheet_id,
                                                'startRowIndex': 3, # Row # 1
                                                'endRowIndex':4, # Row # 10
                                                'startColumnIndex': 2, # column B
                                                'endColumnIndex': 6
                                            }
                                        ]
                                    }
                                },
                                # Y axis values
                                'series': {
                                    'sourceRange': {
                                        'sources': [
                                            {
                                                'sheetId': sheet_id,
                                                'startRowIndex': 2, # Row # 1
                                                'endRowIndex': 3, # Row # 10
                                                'startColumnIndex': 2, # column B
                                                'endColumnIndex': 6
                                            }
                                        ]
                                    }
                                },
                                'bubbleLabels': {
                                    'sourceRange': {
                                        'sources': [
                                            {
                                                'sheetId': sheet_id,
                                                'startRowIndex': 1, # Row # 1
                                                'endRowIndex': 2, # Row # 10
                                                'startColumnIndex': 2, # column B
                                                'endColumnIndex': 6
                                            }
                                        ]
                                    }
                                },
                                'bubbleOpacity': 0.75,
                                'bubbleBorderColorStyle': {
                                    'rgbColor': {
                                        'red': 0,
                                        'green': 0,
                                        'blue': 0,
                                        'alpha': 1
                                    }
                                },
                                'bubbleMinRadiusSize': 20,
                                'bubbleMaxRadiusSize': 20,
                                "bubbleTextStyle": {
                                    "foregroundColor": {
                                        'red':1,
                                        'green':1,
                                        'blue':1,
                                        'alpha':1
                                    },
                                    "fontFamily": 'Comic Sans',
                                    "fontSize": 12,
                                    "bold": True
                                }
                            }
                        },
                        'position': {
                            'overlayPosition': {
                                'anchorCell': {
                                    'sheetId': sheet_id,
                                    'rowIndex': 1,
                                    'columnIndex': 1
                                },
                                'offsetXPixels': 0,
                                'offsetYPixels': 107,
                                'widthPixels': 510,
                                'heightPixels': 400
                            }
                        }
                    }
                }
            }
        ]
    }
    chart_prop = service.spreadsheets().batchUpdate(
            spreadsheetId = file_id,
            body = request_body_bubble
    ).execute()
    chartidnew = chart_prop['replies'][0]['addChart']['chart']['chartId']
    return chartidnew
