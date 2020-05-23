from oauth2client import file, client, tools
import argparse
from googleapiclient import discovery
from apiclient.discovery import build
from pprint import pprint
from httplib2 import Http
from os import path
import sys


TOKEN_PATH_HELP = 'Token path. See `python {} generate_token --help`.'.format(sys.argv[0])
# Brief research suggests that Google Sheets might not be able to have more than 5 million cells
MAX_CELLS_IN_GOOGLE_SHEET = 10_000_000


def generate_token(credentials_path, token_output_path):
    if path.exists(token_output_path):
        print('File ' + token_output_path + ' already exists! Exiting without generating token.')
        sys.exit(1)
    store = file.Storage(token_output_path)
    flow = client.flow_from_clientsecrets(credentials_path, 'https://www.googleapis.com/auth/drive.file')
    tools.run_flow(flow, store, flags=argparse.Namespace(noauth_local_webserver=True, logging_level='ERROR'))


def _get_creds(token_path):
    if not path.exists(token_path):
        print('File ' + token_path + ' does not exist')
    store = file.Storage(token_path)
    creds = store.get()
    if not creds or creds.invalid:
        print('Token at ' + token_path + ' is not valid. Do you have permission to read the file?')
    return creds


def create_sheet(token_path):
    creds = _get_creds(token_path)
    service = discovery.build('sheets', 'v4', credentials=creds)
    spreadsheet_body = {}
    request = service.spreadsheets().create(body=spreadsheet_body)
    return request.execute()


def read_rows(token_path, spreadsheet_id, start_at_row):
    creds = _get_creds(token_path)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    range_name = 'Sheet1!A' + str(start_at_row) + ':' + str(int(start_at_row) + MAX_CELLS_IN_GOOGLE_SHEET)
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    return values


def append_row(token_path, spreadsheet_id, row):
    assert type(row) is list
    row = list(map(str, row))
    creds = _get_creds(token_path)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    body = {'values': [row]}
    return service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, body=body, range='Sheet1!A:B', valueInputOption='RAW').execute()


def _generate_token(args):
    generate_token(args.credentials_path, args.token_path)


def _create_sheet(args):
    sheet_info = create_sheet(args.token_path)
    print('spreadsheetId: ' + sheet_info['spreadsheetId'])
    print('spreadsheetUrl: ' + sheet_info['spreadsheetUrl'])


def _read_rows(args):
    print(read_rows(args.token_path, args.spreadsheet_id, args.start_at_row))


def _append_row(args):
    print(append_row(args.token_path, args.spreadsheet_id, args.entry))


parser = argparse.ArgumentParser(description='Create, append to, and read from Google Sheets that are associated with a Google APIs project, without exposing access to any other Google Sheets or Google Drive data.')
parser.set_defaults(func=lambda x: None)
subparsers = parser.add_subparsers()

generate_token_parser = subparsers.add_parser(
        'generate_token',
        help='Generate a token for accessing Google Drive files via a Google APIs project')
generate_token_parser.set_defaults(func=_generate_token)
generate_token_parser.add_argument('credentials_path', type=str, help='File path to credentials.json, which you can download after creating a project in https://console.developers.google.com/apis/. Enable the Google Sheets API in that project, then create credentials that use OAuth client ID, and download the credentials.')
generate_token_parser.add_argument('token_path', type=str, help='File path to use for writing token.json. Depending on your application, you may need to chmod this token.json file afterwards.')

create_sheet_parser = subparsers.add_parser('create_sheet', help='Create a sheet')
create_sheet_parser.set_defaults(func=_create_sheet)
create_sheet_parser.add_argument('token_path', type=str, help=TOKEN_PATH_HELP)

read_rows_parser = subparsers.add_parser(
        'read_rows',
        help='Read rows from a sheet')
read_rows_parser.set_defaults(func=_read_rows)
read_rows_parser.add_argument('token_path', type=str, help=TOKEN_PATH_HELP)
read_rows_parser.add_argument('spreadsheet_id', type=str, help='Spreadsheet ID')
read_rows_parser.add_argument('start_at_row', type=int, help='Start reading from this one-indexed row number')

append_row_parser = subparsers.add_parser(
        'append_row',
        help='Append a row to a sheet')
append_row_parser.add_argument('token_path', type=str, help=TOKEN_PATH_HELP)
append_row_parser.add_argument('spreadsheet_id', type=str, help='Spreadsheet ID')
append_row_parser.add_argument('entry', nargs='*', help='Entry in the row')
append_row_parser.set_defaults(func=_append_row)

args = parser.parse_args()
args.func(args)
