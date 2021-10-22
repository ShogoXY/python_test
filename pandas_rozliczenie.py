# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import docx
from docx import Document
from docxtpl import DocxTemplate, InlineImage
import numpy as np
from openpyxl import load_workbook
import datetime
import time
from re import search

# file path
path_file = "C:\\Users\\Dariusz\\github\\python_test\\serwi.xlsx"
path_docx = "C:\\Users\\Dariusz\\github\\python_test\\Rozliczenie serwisowe.docx"

doc = docx.Document(path_docx)


ct=(time.strftime('%d.%m.%Y'))
print (ct)
cr = time.strftime('%y%m%d%H%M')
cr2=("CRS"+cr)
print(cr2)

nazwa = input("podaj klienta \n")


doc=DocxTemplate(path_docx)
context = {'date' : ct, 'cr_number' : cr2, 'nazwa' : nazwa, 'data' : ct}
doc.render(context)

# search value in excel sheet (if yes, pritn it)
df = pd.read_excel(path_file, sheet_name='Arkusz1')
print (df)
search_word = input("podaj numer seryjny do wyszukania \n")
search_word2 = str(search_word)
print(type(search_word2))

def search(keyword, df):
    search = '|'.join(keyword)
    searched = df[df['RMA'].str.contains(search, na=False)]
    return searched
    
search_df = search([search_word], df)

excel_book = load_workbook(path_file)
writer = pd.ExcelWriter(path_file, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)


df1 = pd.DataFrame(search_df, columns=['RMA', 'Nazwa urządzenia', 'Nr seryjny przyjęty', 'Nr seryjny wydany', 'UWAGI'])
df3 = pd.DataFrame(search_df, columns=['RMA'])
df1.to_excel(writer, sheet_name='arkusz_test')
print (df1)
# save nd exit excel
writer.save()
writer.close()


for i in range(df1.shape[0]):
    doc.tables[0].add_row() 
    for j in range(df1.shape[-1]):
        table2=doc.tables[0]
        table2.cell(i+1, j+1).text=str(df1.values[i, j])
        table2.cell(i+1, 0).text = str(i+1)



doc.save("rozliczenie "+ cr2 + " " + nazwa+".docx")

