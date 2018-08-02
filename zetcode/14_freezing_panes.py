#!/usr/bin/env python3

from openpyxl import Workbook
from openpyxl.styles import Alignment

book = Workbook()
sheet = book.active

sheet.freeze_panes= 'B2' # congela a celula anterior, A1

book.save('freeze_panes.xlsx')
print('Done.')