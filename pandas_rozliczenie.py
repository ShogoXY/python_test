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
    
    
    
#search_df = df.loc[df["RMA"] == search_word2]

# print value in specific column and search with this value
#komentarz_search = search_df['KOMENTARZ'].values[0]
#print(komentarz_search)
#komentarz_df = df.loc[df['KOMENTARZ'] == komentarz_search]
#print(komentarz_df)

# write value to new sheet in same workbook
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


# word create table in word
table = doc.add_table(df1.shape[0]+1, df1.shape[1], style='Table Grid')

for j in range(df1.shape[-1]):
    table.cell(0, j).text = df1.columns[j]

for i in range(df1.shape[0]):
    for j in range(df1.shape[-1]):
        table.cell(i+1, j).text = str(df1.values[i, j])
        
        
n=i+1
print (n)
for k in range(1,n):
    m = str(k) 
    print("test_" + m +"")        
    table2=doc.tables[0]
    table2.cell(k,1).text=df3   
    doc.tables[0].add_row()     
doc.save(path_docx)


# '
#input = Document('OutputDoc.docx')

#paragraphs = []
# for para in input.paragraphs:
#    p = para.text
#    paragraphs.append(p)

#output = Document()
# for item in paragraphs:
#    output.add_paragraph(item)
# output.save('word1.docx')
#

# Imports

#input_doc = Document('OutputDoc.docx')
#output_doc = Document()

# Call the function


# def get_para_data(output_doc_name, paragraph):
# 
#     output_para = output_doc_name.add_paragraph()
#     for run in paragraph.runs:
#         output_run = output_para.add_run(run.text)
#         output_run.bold = run.bold
#         output_run.italic = run.italic
#         output_run.underline = run.underline
#         output_run.font.color.rgb = run.font.color.rgb
#         output_run.style.name = run.style.name
#         output_run.font.size = run.font.size
#     output_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
#     output_para.paragraph_format.line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
#     output_para.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
#     output_para.paragraph_format.right_indent = paragraph.paragraph_format.left_indent
#     output_para.paragraph_format.first_line_indent = paragraph.paragraph_format.left_indent
# 
# 
# for para in input_doc.paragraphs:
#     get_para_data(output_doc, para)
# 
# output_doc.save('OutputDoc2.docx')
