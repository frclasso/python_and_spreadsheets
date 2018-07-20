#!/usr/bin/env python3

import openpyxl

wb = openpyxl.load_workbook('example.xlsx')
print(wb.get_sheet_names()) # ['Sheet1']

sheet = wb.get_sheet_by_name('Sheet1')
print(sheet)  #  <Worksheet "Sheet1">

print(type(sheet))  # <class 'openpyxl.worksheet.worksheet.Worksheet'>

print(sheet.title) # Sheet1

abaAtiva = wb.get_active_sheet()
print(abaAtiva) #  <Worksheet "Sheet1">, Sheet1 foi a ultima aba visitada