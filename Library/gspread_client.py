#!/usr/bin/env python3
''' Exporting csv file to google spreadsheets '''

from re import search
import re
import sys
import gspread
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from library.fetch_ckms import store_keys



# To generate authentication credentials. See
# https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
#
# Authorize using one of the following scopes:
#     'https://www.googleapis.com/auth/drive'
#     'https://www.googleapis.com/auth/drive.file'
#     'https://www.googleapis.com/auth/spreadsheets'



SCOPE = ('https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive')
CREDENTIALS = ServiceAccountCredentials.from_json_keyfile_name('Media-Hardware.json',
                                                                   SCOPE)
SERVICE = discovery.build('sheets', 'v4', credentials=CREDENTIALS)



def init_cred_gspread_client():
    """
    Initialisng gspread Credentials
    """
    _client = gspread.authorize(CREDENTIALS)
    return _client

def copy_to_csv(sheetname, filename):
    """
    Creating a dummy file
    importing csv to dummy spreadsheet
    """
    CLIENT = init_cred_gspread_client()
    web = CLIENT.create(sheetname)
    web_id = web.id
    web_url = 'https://docs.google.com/spreadsheets/d/{0}'.format(web_id)
    #print(web_url)
    regex = re.search(r'/d/(.*)', web_url)
    new_sheet_id = regex.group(1)
    content = open(filename, 'r').read()
    CLIENT.import_csv(new_sheet_id, content)
    return new_sheet_id

def get_sheetid(SPREADSHEET_ID, ranges, include_grid_data):
    """
    To get the new sheet id where we copy our csv file
    """
    request = SERVICE.spreadsheets().get(spreadsheetId=SPREADSHEET_ID,
                                         ranges=[],
                                         includeGridData=False)
    response = request.execute()
    get_id = response["sheets"][0]["properties"]["sheetId"]
    return get_id

def copy_to_main_spreadsheet(_spreadsheetid, SPREADSHEET_ID, SHEET_ID):
    """
    Copying to spreadsheet to another
    """
    #The ID of the spreadsheet to copy the sheet to.
    copy_sheet_to_another_spreadsheet = {
        'destination_spreadsheet_id': _spreadsheetid,
    }
    request = SERVICE.spreadsheets().sheets().copyTo(spreadsheetId=SPREADSHEET_ID,
                                                     sheetId=SHEET_ID,
                                                     body=copy_sheet_to_another_spreadsheet)
    request.execute()

def get_values(sheetname, SPREADSHEET_ID):
    result = SERVICE.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=sheetname).execute()
    values = result.get('values', [])
    return values

def get_dest_sheetid(sheetname, spreadsheet_id):
    '''Get Destination Sheet Id'''
    request = SERVICE.spreadsheets().get(spreadsheetId=spreadsheet_id,
                                         ranges=[],
                                         includeGridData=False)
    response = request.execute()
    #get_id = response["sheets"][0]["properties"]["sheetId"]

    print("Sheetname = %s" %sheetname)
    print("spreadsheetid = %s" %spreadsheet_id)

    for i in response["sheets"]:
        if search(sheetname, i["properties"]["title"]):
            main_sheet_id = i["properties"]["sheetId"]
            print("Mainsheetid %s" %main_sheet_id)
            return main_sheet_id


def remove_existing_sheet(sheetname, spreadsheet_id):
    request1 = SERVICE.spreadsheets().get(spreadsheetId=spreadsheet_id,
                                          ranges=[],
                                          includeGridData=False)
    response = request1.execute()

    for i in response["sheets"]:
        if search(sheetname, i["properties"]["title"]):
            sheet_id = i["properties"]["sheetId"]
            batch_update_spreadsheet_request_body = {
                "requests": [
                    {
                        "deleteSheet": {
                            "sheetId": sheet_id
                        }
                    }
                ]
            }
            #print(sheet_id)
            request = SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                         body=batch_update_spreadsheet_request_body)
            request.execute()

def copy_to_main_sheet(sheetname, values, spreadsheet_id):
    """
    Append data to main sheet
    """
    range_ = sheetname
    value_input_option = "RAW"
    insert_data_option = "INSERT_ROWS"
    body = {
        'values': values
    }
    #print(body)
    request = SERVICE.spreadsheets().values().append(spreadsheetId=spreadsheet_id,
                                                     range=range_,
                                                     valueInputOption=value_input_option,
                                                     insertDataOption=insert_data_option,
                                                     body=body)
    response = request.execute()

def _batch(self, requests, spreadsheet_id):
    body = {
        'requests': requests
    }
    return self.SERVICE.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                              body=body).execute()

def renameSheet1(self, sheetId, newName):
    return self._batch({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheetId,
                "title": newName,
            },
            "fields": "title",
        }
    })

def renameSheet(sheetId, newName, spreadsheet_id):
    body = {
        "requests": [{
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheetId,
                    "title": newName,
                },
                "fields": "title"
            }
        }]
    }

    return SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                       body=body).execute()

def create_new_sheet(email, sheetname):
    """
    Creating sheet and sharing
    """
    client = init_cred_gspread_client()
    wb = client.create(sheetname)
    wb_id = wb.id
    web_url = 'https://docs.google.com/spreadsheets/d/{0}'.format(wb_id)
    sheetid = re.search(r'/d/(.*)', web_url)
    #onlyid = sheetid.group(1)
    #content = open('test.csv', 'r').read()
    #client.import_csv(onlyid, content)
    wb.share(email, perm_type='user', role='writer')
    print("New Spreadsheet Generated !!!! - %s " %(web_url))
    return wb_id

def add_sheet(sheet_name, spreadsheet_id):
    '''
    Adding new sheet
    '''
    try:
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                    }
                }
            }]
        }

        response = SERVICE.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=request_body
        ).execute()

        return response
    except Exception as e:
        print(e)

def  find_sheet(sheetname, spreadsheet_id):
    '''
    find if already given sheet is existing
    '''
    make_request = SERVICE.spreadsheets().get(spreadsheetId=spreadsheet_id,
                                              ranges=[],
                                              includeGridData=False)
    get_response = make_request.execute()
    for i in get_response["sheets"]:
        if search(sheetname, i["properties"]["title"]):
            sheet_id = i["properties"]["sheetId"]
            return sheet_id
