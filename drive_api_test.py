'''
This file contains the main code
    - uses google spreadhseet api instance
    - takes in input for ticker and days
    - puts values of prices in sheet1
    - puts returns in sheet2
    - puts stats in sheet 3
    - puts graph and data for each set of weightages in sheet4
'''

from Google import Create_Service
import pandas as pd
# from ids import file_id
from ids import *
from functions import createstock,formatCells,sheetclear
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

symbolin1 = input('Ticker Input: ')
symbolin2 = input('Ticker Input: ')

symbols = {
    "symbol01" : symbolin1, 
    "symbol02" : symbolin2
}
days = int(round(int(input("Working days: "))/5,0))
skiprows = 0




# =============================================================================
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
                                "red": 1,
                                "green": 1,
                                "blue": 1,
                                "alpha": 0.3
                            },
                            "type":"MIN" 
                        },
                        "midpoint": {
                            "color": {
                                "red": 0.6,
                                "green": 1,
                                "blue": 0.6,
                                "alpha": 1
                            },
                            "type": "PERCENT" ,
                            "value": "80"
                        },
                        "maxpoint": {
                            "color": {
                                "red": 0,
                                "green": 1,
                                "blue": 0,  
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
response_date = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId = file_id,
        body = request_body_cond
).execute()


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

# createstock(symbol, sheet, columnIndex, file_id, service,sheet2_id,sheet3_id,days)
j = 1
for i in symbols.values():
    createstock(i, sheet01_name, sheet02_name, sheet03_name, j, file_id, sheet_service, sheet02_id, sheet03_id, days)
    j += 1


# =============================================================================

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

# chart_id = chart_prop['replies'][0]['addChart']['chart']['chartId']
# =============================================================================

# =============================================================================
# End of program. Should I clear the sheet?
clear_sheet = input("Clear Sheet? ")
if clear_sheet == 'y':
    sheetclear(sheet_service)
# =============================================================================