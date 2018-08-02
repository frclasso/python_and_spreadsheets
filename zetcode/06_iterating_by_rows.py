#!/usr/bin/env python3

from openpyxl import Workbook
book = Workbook()
sheet = book.active

rows = (
    (88, 46, 57),
    (89, 38, 12),
    (23, 59, 78),
    (56, 91, 28),
    (24, 18,43),
    (34, 15, 67)
)

for row in rows:
    sheet.append(row)

for row in sheet.iter_rows(min_row=1, min_col=1, max_row=6, max_col=3):
    for cell in row:
        print(cell.value, end=' ')
    print()


book.save('iter_row.xlsx')
print('Done')