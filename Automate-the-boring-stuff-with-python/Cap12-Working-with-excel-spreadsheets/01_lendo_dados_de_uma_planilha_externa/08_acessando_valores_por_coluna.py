#!/usr/bin/env python3

import openpyxl

import openpyxl
wb = openpyxl.load_workbook("example.xlsx")
sheet = wb.get_active_sheet()
sheet.columns[1] # TypeError: 'generator' object is not subscriptable
