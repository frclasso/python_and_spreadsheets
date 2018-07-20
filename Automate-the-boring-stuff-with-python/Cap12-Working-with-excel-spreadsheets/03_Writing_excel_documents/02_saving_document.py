#!/usr/bin/env python3

import openpyxl

wb = openpyxl.Workbook()
#print(wb.get_sheet_names())
sheet = wb.get_active_sheet()
sheet.title = 'Folha1'
print(wb.get_sheet_names())
wb.save('exemplo2.xlsx')