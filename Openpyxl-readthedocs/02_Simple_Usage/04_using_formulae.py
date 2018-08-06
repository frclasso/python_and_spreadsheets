#!/usr/bin/env python3

from openpyxl import Workbook
wb = Workbook()
ws = wb.active

ws["A1"] = "=SUM(1, 1)"
wb.save('formula.xlsx')
print('Done.')