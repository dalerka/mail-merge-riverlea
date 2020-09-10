# -*- coding: utf-8 -*-
#
# Copyright ©2018-2019 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at apache.org/licenses/LICENSE-2.0.
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
docs-mail-merge.py (Python 2.x or 3.x)

Google Docs (REST) API mail-merge sample app

https://developers.google.com/docs/api/samples/mail-merge
"""
# [START mail_merge_python]
from __future__ import print_function

import math
import time
import json

from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

# Fill-in IDs of your Docs template & any Sheets data source
#The Sheet Tempalte
DOCS_FILE_ID = '1wKnXBK7x2r3YKbQ-dBC9WgGYk69pnQ3VkSBh4lU0LUg'
#The Data
SHEETS_FILE_ID = '1pnM4Ndq1Q8OQD6_wUHqFauicJtHwLPT4XivdZz7o63g'

# authorization constants
CLIENT_ID_FILE = 'credentials.json'
TOKEN_STORE_FILE = 'token.json'
SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
)

# Not sure exactly how to elimate this
SOURCES = ('text', 'sheets')

template_page_setup = {
    'title': "Created form letter for Riverelea Mail Merge",
    "documentStyle": {
        "marginTop": {
          "magnitude": 18,
          "unit": "PT"
        },
        "marginLeft": {
          "magnitude": 18,
          "unit": "PT"
        },
        "marginRight": {
          "magnitude": 18,
          "unit": "PT"
        }
      }
}

table_columns=3
table_rows=10
table_cell_total=table_rows*table_columns

def get_http_client():
    """Uses project credentials in CLIENT_ID_FILE along with requested OAuth2
        scopes for authorization, and caches API tokens in TOKEN_STORE_FILE.
    """
    store = file.Storage(TOKEN_STORE_FILE)
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_ID_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    return creds.authorize(Http())

# service endpoints to Google APIs
HTTP = get_http_client()
DRIVE = discovery.build('drive', 'v3', http=HTTP)
DOCS = discovery.build('docs', 'v1', http=HTTP)
SHEETS = discovery.build('sheets', 'v4', http=HTTP)

def get_data():
    return SAFE_DISPATCH['sheets']()


def _get_sheets_data(service=SHEETS):
    """(private) Returns data from Google Sheets source. It gets all rows of
        'Sheet1' (the default Sheet in a new spreadsheet), but drops the first
        (header) row. Use any desired data range (in standard A1 notation).
    """

    data_sheet = service.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID,range='Sheet1').execute()

    sheet_header = data_sheet.get('values')[0]
    sheet_data = data_sheet.get('values')[1:] # skip header row

    return [sheet_header, sheet_data]


# data source dispatch table [better alternative vs. eval()]
SAFE_DISPATCH = {k: globals().get('_get_%s_data' % k) for k in SOURCES}

def _create_template( service):
    """Creates an Empty Document.
    """

    return DOCS.documents().create(body=template_page_setup).execute().get('documentId')

def _copy_template(tmpl_id, service):
    """(private) Copies letter template document using Drive API then
        returns file ID of (new) copy.
    """
    body = {'name': 'Merged form letter (%s)' % "sheet"}
    return service.files().copy(body=body, fileId=tmpl_id, fields='id').execute().get('id')

def create_label_table( document_id, service):
    requests = [{
        'insertTable': {
            'rows': table_rows,
            'columns': table_columns,
            'endOfSegmentLocation': {
                'segmentId': ''
            }
        },
    }
    ]
    result = DOCS.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

    return result

def insert_text( document_id, data):

    s1 = """
    {}
    {}
    or current resident
    {}
    Riverlea, OH 43085
    """

    requests = []
#We must go backwards

    index = table_cell_total

    if len(data) < table_cell_total:
        index = len(data)

    while (index > 0):
        index -= 1

        row = data[index]
        merged = s1.format( row[0], row[1],row[2])

        # For each index, determine row and column
        cur_row = math.trunc( index / table_columns)
        cur_col =  index % table_columns
        table_index = ( cur_row * 7 ) + 5 + ( 2 * cur_col )
        requests.append( {'insertText': {'text': merged, 'location': {'index': table_index}} } )

    result = DOCS.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

    return result

def create_template(tmpl_id, service):
    """Copies template document and merges data into newly-minted copy then
        returns its file ID.
    """
    # copy template and set context data struct for merging template values
    destination_id = _create_template( service)

    create_label_table(destination_id, service)

    return destination_id

def do_mail_merge(tmpl_id, data ):

    document_id = create_template(tmpl_id, DRIVE)

    insert_text( document_id, data)

    return document_id

def chunks(l, n):
    n = max(1, n)
    return (l[i:i+n] for i in range(0, len(l), n))

if __name__ == '__main__':
 #   # get row data, then loop through & process each form letter
    columns, data = get_data() # get data from data source

    data_rows = len( data )
    num_sheets = data_rows / table_cell_total

    for i, datablock in enumerate( chunks( data, table_cell_total )):
        document_id = do_mail_merge( DOCS_FILE_ID, datablock )
        print('Merged letter: docs.google.com/document/d/%s/edit' % document_id )
        i+1
# [END mail_merge_python]