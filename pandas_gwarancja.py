# bedziemy próbować zrobić działający program na plikach excel za pomocą Pandas
import pandas as pd
import numpy as np
import openpyxl
#wb = openpyxl.load_workbook('/home/darek/python/test.xlsx')
df = pd.read_excel('/home/darek/python/test.xlsx')

print(df.head())

search_word = input()

ppd = df.loc[df['nazwa'] == search_word]

s_ddp = ppd['cena'].values[0]
print(s_ddp)
p_ddp = df.loc[df['cena'] == s_ddp]
print(p_ddp)


df1 = pd.DataFrame(p_ddp)
df1.to_excel("/home/darek/python/test.xlsx", sheet_name='test2')

# wb.save('/home/darek/python/test.xlsx')
