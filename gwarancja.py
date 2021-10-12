# program do tworzenia karty gwarancyjnej w PDF na podstawie danych z Excel
import os
# import pandas as pd
# import docxptl
# import docx
import openpyxl
import sys

os.chdir("/home/darek/git/python_test")
wb = openpyxl.load_workbook('test.xlsx')
ws = wb['Arkusz1']
# sn_warrany=input("Podaj numer seryjny do druku gwarancji\n")
if 'dane' not in wb.sheetnames:
    wb.create_sheet('dane')
ws_1 = wb['dane']

searchString = input("podaj wartość którą chcesz wyszukać\n")
for i in range(1, ws.max_row + 1):
    if searchString == ws.cell(i, 5).value:
        commet = ws.cell(i, 11).value
        print(commet)
        for i in range(1, ws.max_row + 1):
            if commet == ws.cell(i, 11).value:
                pn = ws.cell(i, 4).value
                sn = ws.cell(i, 5).value
                gwara1 = str(pn)
                gwara2 = str(sn)
                full_gwara = (gwara1 + ", " + gwara2)
                print(full_gwara)
            for j in range(1, ws_1.max_row + 1):
                ws_1.cell(j + 1, 1).value = gwara1
                ws_1.cell(j + 1, 2).value = gwara2

            #    ws_1.cell(j, 2).value = sn
# print(len(gwara))
# Save the openpyxl Workbook object to file
wb.save('test.xlsx')
