#!/usr/bin/env python3

import openpyxl
wb = openpyxl.Workbook()

ws = wb.create_sheet()
ws.column_dimensions.group('A','D',hidden=True)  # Esconde celulas definidas no intervalo
wb.save('group.xlsx')
print('Done.')