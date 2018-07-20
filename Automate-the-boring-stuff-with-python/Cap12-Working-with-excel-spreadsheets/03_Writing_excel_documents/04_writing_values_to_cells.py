#!/usr/bin/env python3

import openpyxl

"""Writing values to cells is much like writing values to keys in a dictionary"""

wb = openpyxl.Workbook()
sheet = wb.get_sheet_by_name('Sheet')
sheet['A1'] = 'Hello Python'
print(sheet['A1'].value)