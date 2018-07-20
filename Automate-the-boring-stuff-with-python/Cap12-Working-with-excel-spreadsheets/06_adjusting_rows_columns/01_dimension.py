#!/usr/bin/env python3

import openpyxl

wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
sheet["A1"] = "Tall row"
sheet["B2"] = "Wide column"
sheet.row_dimensions[1].height = 70 # Altura da linha
sheet.column_dimensions['B'].width = 50  # Largura da coluna

wb.save('dimensionsSheet.xlsx')
print('Done')