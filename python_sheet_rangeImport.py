#!/usr/bin/env python
# pylint: disable=no-member

# Author: Felix Arias
# Date 1/12/2021

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes or changing account, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# ************SHEET CONFIGURATIONS********************** #
# Name: Sample Source Sheet
# Description: Pulling from this sheet.
Sheet1 = {'url' :'1u9j3vf0p4ORnvx7ozql_6WORNfGn-zho1CEJBWC1a7s', 
'range':'Data Pull!A:E',
'range2':'Data Pull!F:G'}

# Name: Test
# Description: Spreadsheet made to receive a copy of the range in source.
Sheet2 = {'url' :'1nydDconQhxpyOUDKm2rh0i8NAbk7RwAunSYxk4sjaVw',
'range':'Sheet1!A:E',
'range2':'Sheet1!F:E'}

# Main function
def main():
    """Shows basic usage of the Sheets API.Prints values from a sample spreadsheet.
    """
    creds = credentials()
    importRange(Sheet1['url'],Sheet1['range'],Sheet2['url'],Sheet2['range'],creds)
    importRange(Sheet1['url'],Sheet1['range2'],Sheet2['url'],Sheet2['range2'],creds)

# Reads from a Sheet
def readSheet(SPREADSHEET_SOURCE,SOURCE_RANGE_NAME,creds):
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    print('Reading...')
    result = sheet.values().get(spreadsheetId=SPREADSHEET_SOURCE, range=SOURCE_RANGE_NAME).execute()
    values = result.get('values', [])
    return values

# Writes to a Sheet
def writeSheet(SPREADSHEET_DESTINATION,DESTINATION_RANGE_NAME,values,creds):
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    value_input_option = 'RAW'  # RAW makes it not try to calculate or convert number types.
    print('Writing...')
    body = {'values' : values}
    request = sheet.values().update(spreadsheetId=SPREADSHEET_DESTINATION, range=DESTINATION_RANGE_NAME, valueInputOption=value_input_option, body=body)
    response = request.execute()
    
    print("This is the link to the destination:\n https://docs.google.com/spreadsheets/d/%s/edit" % response['spreadsheetId'])
    return response

# Creates Credentials from user's credentials.json file.
def credentials():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

# Import Range function
def importRange(SPREADSHEET_SOURCE,SOURCE_RANGE_NAME,SPREADSHEET_DESTINATION,DESTINATION_RANGE_NAME,creds):    
    values = readSheet(SPREADSHEET_SOURCE,SOURCE_RANGE_NAME,creds)    
    if not values:  # Check if Range is empty
        print('No data found(Source). Ending.')
        exit()  # Quit if nothing was found.
    else:
        print(writeSheet(SPREADSHEET_DESTINATION,DESTINATION_RANGE_NAME,values,creds))


if __name__ == '__main__':
    main()
