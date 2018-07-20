#!/usr/bin/env python3

"""Podemos fatiar obejetos de uma Woksheets para obtermos Cell objects de linhas
 colunas ou uma area retangular da planilha(sheet)"""

import openpyxl
wb = openpyxl.load_workbook("example.xlsx")
sheet = wb.get_sheet_by_name('Sheet1')
# Obtendo Cell objects
print(tuple(sheet['A1': 'C3']))

print()
# Interando sobre os objetos
for rowOfCellObjects in sheet['A1':'C3']:
    for cellObject in rowOfCellObjects:
        print(cellObject.coordinate, cellObject.value)
    print("----- END OF ROW -----")
    print()

