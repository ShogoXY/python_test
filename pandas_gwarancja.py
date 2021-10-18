# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import numpy as np
from openpyxl import load_workbook

file = "/home/fedora/git/python_test/test.xlsx"

df = pd.read_excel(file, sheet_name='Produkcja')

print(df.head())

search_word = input("podaj numer seryjny do wyszukania \n")

ppd = df.loc[df['NUMER SERYJNY'] == search_word]

s_ddp = ppd['KOMENTARZ'].values[0]
print(s_ddp)
p_ddp = df.loc[df['KOMENTARZ'] == s_ddp]
print(p_ddp)


excel_book = load_workbook(file)
writer = pd.ExcelWriter(file, engine='openpyxl',mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)

df1 = pd.DataFrame(p_ddp, columns=['NAZWA', 'NUMER SERYJNY'])
df1.to_excel(writer, sheet_name='arkusz')
writer.save()
writer.close()
