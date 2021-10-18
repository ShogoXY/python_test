# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import docx
import numpy as np
from openpyxl import load_workbook

# ścieżka dostępu do pliku
path_file = "/home/fedora/git/python_test/test.xlsx"
path_docx = "/home/fedora/git/python_test/word1.docx"

doc = docx.Document(path_docx)

# wyszukuje wartości w akruszu (czy jest, jeśli tak wypisuje ją)
df = pd.read_excel(path_file, sheet_name='Produkcja')
search_word = input("podaj numer seryjny do wyszukania \n")
search_df = df.loc[df['NUMER SERYJNY'] == search_word]

# wypisuje warość przypisaną do szukanej a następnie wypisuje na tej podstawie tabelę
komentarz_search = search_df['KOMENTARZ'].values[0]
print(komentarz_search)
komentarz_df = df.loc[df['KOMENTARZ'] == komentarz_search]
print(komentarz_df)

# wpisuje wyszukiwane wartości do nowego arkusza, nadpisując je
excel_book = load_workbook(path_file)
writer = pd.ExcelWriter(path_file, engine='openpyxl',
                        mode='a', if_sheet_exists='replace')
writer.book = excel_book
writer.sheets = dict((ws.title, ws) for ws in excel_book.worksheets)
df1 = pd.DataFrame(komentarz_df, columns=['NAZWA', 'NUMER SERYJNY'])
df1.to_excel(writer, sheet_name='arkusz')

# zapisz wyjdź
writer.save()
writer.close()

t = doc.add_table(df1.shape[0]+1, df1.shape[1])

for j in range(df1.shape[-1]):
    t.cell(0, j).text = df1.columns[j]

for i in range(df1.shape[0]):
    for j in range(df1.shape[-1]):
        t.cell(i+1, j).text = str(df1.values[i, j])

doc.save(path_docx)
