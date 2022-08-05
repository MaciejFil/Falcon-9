import pandas as pd
import openpyxl as op
import re
import datetime
import os

# ŚCIEŻKI DO PLIKÓW
dir_path = os.path.dirname(os.path.realpath(__file__))
path_to_output = dir_path + "/output.xlsx"

# ADRESY URL
url = pd.read_html('https://en.wikipedia.org/wiki/List_of_Falcon_9_and_Falcon_Heavy_launches_(2010%E2%80%932019)')
url2 = pd.read_html('https://en.wikipedia.org/wiki/List_of_Falcon_9_and_Falcon_Heavy_launches')

# KONWERSJA LISTY LIST NA DATEFRAME
url = pd.concat(url[0:9])
url2 = pd.concat(url2[0:5])
df = pd.concat([url, url2])

df.reset_index(inplace=True, drop=True)
new_columns = ['Flight No.','Date and time', 'Booster version', 'Launch site', 'Payload', 'Payload mass', 'Orbit', 'Customer', 'Launch outcome', 'Boosterlanding', 'Description']

# USUWANIE ZBĘDNYCH WIERSZY I KOLUMN
nan_value = float("NaN")
df.replace("", nan_value, inplace=True)
df.dropna(subset = ['Flight No.'], inplace=True)
df.drop([0,1], axis=1, inplace=True)

# WYPEŁNIANIE BRAKÓW W KOLUMNACH
df['Version,Booster [a]'].fillna(df['Version,Booster[a]'], inplace=True)
df['Version,Booster [a]'].fillna(df['Version,booster[b]'], inplace=True)
df['Payload[b]'].fillna(df['Payload[c]'], inplace=True)
df['Launch site'].fillna(df['Launchsite'], inplace=True)

df.drop(['Version,Booster[a]', 'Version,booster[b]', 'Payload[c]', 'Launchsite'], axis=1, inplace=True)

# usuwanie ostatnich duplikatów
duplicated_row = df[df.duplicated(keep='last', subset=['Flight No.'])] # odrzucenie ostatnich duplikatów
df1 = pd.concat([df, duplicated_row]).drop_duplicates(keep=False) # tworzenie roboczej df z ostatnimi duplikatami
df['description'] = df1['Version,Booster [a]'] # dodanie roboczej df z ostatnimi duplikatami do właściwej df
df['description'].fillna(method='backfill', inplace=True) # uzupełnianie braków w kolumnie z opisem

df = df[df.duplicated(keep='last', subset=['Flight No.'])] # usuwanie zbędnych wierszy z ostatnimi duplikatami

df.reset_index(drop=True, inplace=True)

df.replace(r'\[.*\]', '', regex=True, inplace=True)
df['Payload mass'].replace(r'\(.*lb\)', '', regex=True, inplace=True)
df['description'].replace(r'\(more details\)', '', regex=True, inplace=True)

# Zmiana nazw kolumn
df.columns = new_columns

# PORZĄDKOWANIE KOLUMNY Z DATAMI I KONWERSJA NA DATATIME
new_row_list = []
for x in df['Date and time']:
    if ':' in x:
        text_adjustment = re.search(r'\d\d:', x)
        start = text_adjustment.start()
        x = x.replace(x[start:], ' ' + x[start:])
        x = x.replace(',', '')
    x = x.replace('(planned)', '')
    if bool(re.search(r':\d\d:', x)):
        x = x
    elif bool(re.search(r' \d\d:\d\d', x)):
        x = x + ':00'
    else:
        x = x + ' 00:00:00'
    x = datetime.datetime.strptime(x, '%d %B %Y %H:%M:%S')
    print(x)
    new_row_list.append(x)

df['Date and time'] = new_row_list

df.to_excel(path_to_output)

work_book = op.load_workbook(path_to_output, read_only=False)
work_sheet = work_book.active
work_sheet.freeze_panes = 'A2'
work_sheet.auto_filter.ref = 'B1:L1'
work_book.save(path_to_output)

print('Program z radością wykonał Twoje polecenie, Mój Panie!')