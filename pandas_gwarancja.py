# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import docx
from docx import Document
from docxtpl import DocxTemplate, InlineImage
import numpy as np
from openpyxl import load_workbook

# file path
path_file = "C:\\Users\\Dariusz\\github\\python_test\\test.xlsx"
path_docx = "C:\\Users\\Dariusz\\github\\python_test\\word1.docx"

doc = docx.Document(path_docx)

# search value in excel sheet (if yes, pritn it)
df = pd.read_excel(path_file, sheet_name='Produkcja')
search_word = input("podaj numer seryjny do wyszukania \n")
search_df = df.loc[df['NUMER SERYJNY'] == search_word]

# print value in specific column and search with this value
komentarz_search = search_df['KOMENTARZ'].values[0]
print(komentarz_search)
komentarz_df = df.loc[df['KOMENTARZ'] == komentarz_search]
print(komentarz_df)

# write value to new sheet in same workbook
excel_book = load_workbook(path_file)
writer = pd.ExcelWriter(path_file, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)
df1 = pd.DataFrame(komentarz_df, columns=['NAZWA', 'NUMER SERYJNY'])
df1.to_excel(writer, sheet_name='arkusz')

# save nd exit excel
writer.save()
writer.close()


# word create table in word
table = doc.add_table(df1.shape[0]+1, df1.shape[1], style='Table Grid')


for j in range(df1.shape[-1]):
    table.cell(0, j).text = df1.columns[j]

for i in range(df1.shape[0]):
    for j in range(df1.shape[-1]):
        table.cell(i+1, j).text = str(df1.values[i, j])


print(str(df1.values[i, j]))
doc.save(path_docx)


input = Document('word1.docx')

paragraphs = []
for para in input.paragraphs:
    p = para.text
    paragraphs.append(p)

output = Document()
for item in paragraphs:
    output.add_paragraph(item)
output.save('OutputDoc.docx')
