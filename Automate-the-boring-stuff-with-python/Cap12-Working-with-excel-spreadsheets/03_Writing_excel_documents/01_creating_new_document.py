#!/usr/bin/env python3

"""Call the openpyxl.Workbook() function to create a new, blank Workbook object. """

import openpyxl

wb = openpyxl.Workbook() # cria um novo documento tipo Workbook
#print(wb.get_sheet_names()) # ['Sheet']
sheet = wb.get_active_sheet()
#print(sheet.title) # Sheet
sheet.title = 'Spam Bacon Eggs Sheet'
print(wb.get_sheet_names())