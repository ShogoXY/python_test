# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import docx
from docx import Document
from docxtpl import DocxTemplate, InlineImage
import numpy as np
import openpyxl
# from openpyxl import load_workbook

import datetime
import time
from re import search
import keyboard

# file path
path_file = "C:\\Users\\Dariusz\\github\\python_test\\serwi.xlsx"
path_docx = "C:\\Users\\Dariusz\\github\\python_test\\Rozliczenie serwisowe.docx"

doc = docx.Document(path_docx)


ct = (time.strftime('%d.%m.%Y'))
print(ct)
cr = time.strftime('%y%m%d%H%M')
cr2 = ("CRS"+cr)
print(cr2)

nazwa = input("podaj klienta \n")


doc = DocxTemplate(path_docx)
context = {'date': ct, 'cr_number': cr2, 'nazwa': nazwa, 'data': ct}
doc.render(context)

# search value in excel sheet (if yes, pritn it)
df = pd.read_excel(path_file, sheet_name='Arkusz1')
print(df)
# search_word = input("podaj numer RMA \n")
# search_word2 = str(search_word)
# print(type(search_word2))
df_name = "arkusz_test"
excel_book = openpyxl.load_workbook(path_file)
print(excel_book.sheetnames)


if df_name not in excel_book.sheetnames:
    excel_book.create_sheet(df_name)

print(excel_book.sheetnames)
excel_book.save(path_file)

writer = pd.ExcelWriter(path_file, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)

while 1:
    def search_value(keyword, df):
        search_value = '|'.join(keyword)
        searched = df[df['RMA'].str.contains(search_value, na=False)]
        return searched

    search_word = input("podaj numer RMA \n")
    search_df = search_value([search_word], df)

    if search_word == "":
        print("koniec")
        break
    else:

        df1 = pd.DataFrame(search_df, columns=['RMA', 'Nazwa urządzenia',
                           'Nr seryjny przyjęty', 'Nr seryjny wydany', 'UWAGI'])

        # df1.to_excel(writer, sheet_name='arkusz_test')
        df1.to_excel(writer, sheet_name='arkusz_test')

        writer.save()
        df2 = pd.read_excel(path_file, sheet_name='arkusz_test')
        df3 = df2.append(df1)
        # df2 = df1
        # df1.loc[df1.shape[0]]
        # df3.to_excel(writer, sheet_name='arkusz_test')
        print(df3)

        writer.save()

    # save nd exit excel
        continue

writer.save()
writer.close()
#

df5 = pd.DataFrame(['RMA', 'Nazwa urządzenia', 'Nr seryjny przyjęty',
                    'Nr seryjny wydany', 'UWAGI'])
df5.to_excel(writer, sheet_name='arkusz_test')
print(df5)
for i in range(df3.shape[0]):
    doc.tables[0].add_row()
    for j in range(df3.shape[-1]):
        table2 = doc.tables[0]
        table2.cell(i+1, j+1).text = str(df3.values[i, j])
        table2.cell(i+1, 0).text = str(i+1)


doc.save("rozliczenie " + cr2 + " " + nazwa+".docx")
excel_book.remove_sheet("arkusz_test")
