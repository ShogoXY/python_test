# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import numpy as np
from openpyxl import load_workbook

path_file = "/home/fedora/git/python_test/test.xlsx"

df = pd.read_excel(path_file, sheet_name='Produkcja')

print(df.head())

search_word = input("podaj numer seryjny do wyszukania \n")

search_df = df.loc[df['NUMER SERYJNY'] == search_word]

komentarz_search = search_df['KOMENTARZ'].values[0]
print(komentarz_search)
komentarz_df = df.loc[df['KOMENTARZ'] == komentarz_search]
print(komentarz_df)


excel_book = load_workbook(path_file)
writer = pd.ExcelWriter(path_file, engine='openpyxl', mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)

df1 = pd.DataFrame(komentarz_df, columns=['NAZWA', 'NUMER SERYJNY'])
df1.to_excel(writer, sheet_name='arkusz')
writer.save()
writer.close()
