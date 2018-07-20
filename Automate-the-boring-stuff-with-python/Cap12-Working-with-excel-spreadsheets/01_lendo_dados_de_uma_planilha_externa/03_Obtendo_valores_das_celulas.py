#!/usr/bin/env python3

import openpyxl

wb = openpyxl.load_workbook('example.xlsx')
sheet = wb.get_sheet_by_name("Sheet1")
print(sheet['A1']) # <Cell 'Sheet1'.A1>

# Imprimindo os valores das celulas
print(sheet["A1"].value)

b = sheet['B1']
print(b.value)

print("Row " + str(b.row) + ", Column " + b.column + ", is " + b.value)

print("Cell " + b.coordinate + " is " + b.value)

print(sheet['C1'].value)


"""The Cell object has a value attribute that contains, unsurprisingly, the 
value stored in that cell. Cell objects also have row, column, and coordinate attri-
butes that provide location information for the cell."""