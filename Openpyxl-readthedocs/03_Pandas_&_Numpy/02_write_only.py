#!/usr/bin/env python3

from openpyxl.cell.cell import WriteOnlyCell
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook(write_only=True)

ws = wb.create_sheet(title='New')

cell = WriteOnlyCell(ws)
cell.style = 'Pandas'


def format_first_row(row, cell):
    for c in row:
        cell.value = c
        yield cell


rows = dataframe_to_rows(df)
first_row = format_first_row(next(rows),cell)
ws.append(first_row)

for row in rows:
    row = list(row)
    cell.value = row[0]
    row[0] = cell
    ws.append(row)

wb.save('openpyxl_stream.xlsx')