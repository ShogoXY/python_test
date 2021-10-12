# program do tworzenia karty gwarancyjnej w PDF na podstawie danych z Excel
import os
#import pandas as pd
#import docxptl
#import docx
import openpyxl
import sys

os.chdir("/home/fedora/git/python_test")
wb = openpyxl.load_workbook('test.xlsx')
ws = wb.active

#sn_warrany=input("Podaj numer seryjny do druku gwarancji\n")


search_word = input("podaj wartość którą chcesz wyszukać\n")


def wordfinder(searchString):
    for i in range(1, ws.max_row + 1):
        if searchString == ws.cell(i, 5).value:
            commet = ws.cell(i, 11).value
            print(commet)


def copy_to_sh(commet):
    for i in range(1, ws.max_row + 1):
        if commet == ws.cell(i, 11).value:
            pn = ws.cell(i, 4).value
            sn = ws.cell(i, 5).value
            print(sn, pn)


wordfinder(search_word)
