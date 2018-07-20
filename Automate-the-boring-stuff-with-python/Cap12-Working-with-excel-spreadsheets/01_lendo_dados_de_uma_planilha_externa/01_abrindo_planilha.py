#!/usr/bin/env python3

import openpyxl

wb = openpyxl.load_workbook('example.xlsx')
print(type(wb))


"""Nesse exemplo vamos abrir a planilha example.xls, localizada neste mesmo  diretorio
   Workbook = nome do arquivo, no caso example.xls
   Sheet's = planilhas internas (folhas ou abas)

"""