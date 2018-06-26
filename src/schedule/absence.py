'''
Created on Jul 28, 2017

@author: venturf2
'''
import httplib2
import os
from datetime import datetime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_data():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1b_cPFW32a6DTW6uUPR1SgoB7hUNOMsuaLGB2eS_uaHw/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1yllnHABIdM-GLVDfwgwIHDJ8RGM0k_qLSinb3Hy4XeI'
    rangeName = 'Form Responses 1!A:D'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        return values


def create_abs_str(member_name, absence_date):
    absence_str = '\n'.join(
        ['MEMBER_NAME: {member_name}',
         'EXCEPTION_TYPE: 64',
         'DATE: {absence_date}',
         'UNTIL_DATE: {absence_date}',
         'WHICH_DAY_OR_WEEK: 0',
         'END_RECORD',
         '#------------------------------------------------------------',
         ''
         ])
    absence_str = absence_str.format(**{'member_name': member_name,
                                        'absence_date': absence_date})
    return absence_str
    
# MEMBER_NAME: Zhanna Zhilina
# EXCEPTION_TYPE: 64
# DATE: 6-29-2017
# UNTIL_DATE: 7-27-2017
# WHICH_DAY_OR_WEEK: 0
# END_RECORD
# #------------------------------------------------------------
member_dict = {}

def load_members():
    file_path = r'E:\Ventana Drive\Ventana (Work Stuff)\Toastmasters'\
                 '\VentanaVoices\VP Ed. scheduling\TMI\ClubScheduler\Data'\
                 '\TMMembers.dat'
    cur_name = None
    cur_email = None
    with open(file_path, 'r') as member_file:
        for line in member_file:
            if 'NAME:' in line:
                cur_name = line.split('NAME: ')[1].replace('\n', '')
            if 'EMAIL:' in line:
                cur_email = line.split('EMAIL: ')[1].lower().replace('\n', '')
                member_dict[cur_email] = cur_name



def get_name(email):
    if email in member_dict:
        return member_dict[email]
    else:
        print('Missing Email: %s', email)
        raise AttributeError


def get_date(absent):
    '''
    Incoming format: DD-MMM
    Needs to return format M-D-YYYY; why, I don't know
    '''
    absent = absent.rstrip().lstrip()
    try:
        datetime_object = datetime.strptime(absent, '%d-%b')
    except ValueError:
        raise ValueError(f"Date could not be processed: {absent}")
    datetime_object = datetime_object.replace(year=2018)
    out_time = datetime_object.strftime('%m-%d-%Y')
    out_time = out_time.lstrip("0").replace('-0', '-')
    return out_time


def main():
    data = get_data()
    print("FOUND DATA\n", data)
    file_path = r'E:\Ventana Drive\Ventana (Work Stuff)\Toastmasters'\
                 '\VentanaVoices\VP Ed. scheduling\TMI\ClubScheduler\Data'\
                 '\Exceptions.dat'
    with open(file_path, 'a') as abs_file:
        load_members()
        print(member_dict)
        for info in data:
            if len(info) < 3:
                continue
            email = info[1]
            if "Email Address" in email:
                continue
            absent_list = info[2]
            if len(info) > 3:
                func_req = info[3]
            else:
                func_req = None
            member_name = get_name(email)
            for absent in absent_list.split(', '):
                if len(absent) == 0:
                    continue
                absence_date = get_date(absent)
                abs_str = create_abs_str(member_name, absence_date)
                print(abs_str)
                abs_file.write(abs_str)


if __name__ == '__main__':
    #from schedule import absence
    print(True)
    main()
