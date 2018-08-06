#!/usr/bin/env python3

import openpyxl
wb = openpyxl.load_workbook('teste.xlsx')
wb.template = True
wb.save('balances_template.xlsx')  # cria um copia exata do Workbook teste.xlsx