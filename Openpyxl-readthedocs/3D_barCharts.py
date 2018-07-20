#!/usr/bin/env python3

from openpyxl import Workbook
from openpyxl.chart import Reference,Series,BarChart3D

wb = Workbook()
ws = wb.active

rows = [
    ("Produce", 2013, 2014),
    ("Apples", 5, 4),
    ("Oranges", 6, 2),
    ("Pears", 8, 3)
]

for row in rows:
    ws.append(row)

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=4)
titles = Reference(ws, min_col=1, min_row=2, max_row=4)
chart = BarChart3D()
chart.title = "3D Bar Chart"  # Erro: Nao aparece o titulo
chart.add_data(data=data, titles_from_data=True)
chart.set_categories(titles)

ws.add_chart(chart, "E5")
wb.save("bar3d.xlsx")
print("Done.")


"""https://openpyxl.readthedocs.io/en/2.5/charts/bar.html"""
