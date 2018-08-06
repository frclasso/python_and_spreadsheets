#!/usr/bin/env python3

from  openpyxl import load_workbook
wb = load_workbook(filename='empty_book.xlsx')
sheet_range = wb['range names']
print(sheet_range['D18'].value)

sheet_range2 = wb['Pi']
print(sheet_range2['F5'].value)

print('Done')