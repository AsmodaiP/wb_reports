'''Модуль для обновения гугл таблицы'''
from operator import index
import os

from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

import report

load_dotenv()

TELEGRAM_TOKEN = os.environ['TELEGRAM_TOKEN']

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials_service.json')
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)


def convert_to_column_letter(column_number):
    '''Конверертирует номер столбца в его литерал'''
    column_letter = ''
    while column_number != 0:
        c = ((column_number-1) % 26)
        column_letter = chr(c+65)+column_letter
        column_number = (column_number-c)//26
    return column_letter

def update_table(name_of_xlsx='report.xlsx', spreadsheet_id='1cNNK_IPAUt7LAevVaX7mfNudTL7mu4dc14iXZv7TkS8',
                range_name='03.01.-09.01.2022'):
    '''
    Функция служит для обновления таблицы, индексы служат для обозначения куда и какие данные заносить
    '''
    indexes = {
    'артикул': 3,
    'продано': 7,
    'получено от клиента': 12,
    'комиссия': 13,
    'Возвраты': 'AF2'
    }


    relized = report.get_unsorted_relized(name_of_xlsx)
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=range_name, majorDimension='ROWS').execute()

    values = result.get('values', [])

    error_aricles=[]
    body_data = []
    i = 1
    for row in values:
        if len(row)>2:
            try:
                article = row[indexes['артикул']]
                data = relized[article]
                body_data += [
                    {
                        'range':f'{range_name}!{convert_to_column_letter(indexes["продано"])}{i}',
                        'values':[[data['Количество']]]
                    },
                    {
                        'range':f'{range_name}!{convert_to_column_letter(indexes["получено от клиента"])}{i}',
                        'values':[[data['Вайлдберриз реализовал']]]
                    },
                    {
                        'range':f'{range_name}!{convert_to_column_letter(indexes["комиссия"])}{i}',
                        'values':[[data['Вознаграждение Вайлдберриз без НДС']]]
                    },

                ]
                del relized[article]
                i += 1
            except KeyError:
                error_aricles.append(article)
                i += 1
    refund_sums = report.get_sum_of_refunds(name_of_xlsx)
    body_data += [ {'range': f'{range_name}!{indexes["Возвраты"]}', 'values':[[refund_sums]]} ]
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': body_data
    }
    
    sheet.values().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return {'errors': relized, 'sum': (refund_sums)}

if __name__ == '__main__':
    print(update_table()['sum'])
