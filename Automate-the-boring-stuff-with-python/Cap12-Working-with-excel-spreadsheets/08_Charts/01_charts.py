#!/usr/bin/env python3

import openpyxl
from openpyxl.chart import area_chart,BarChart,Series, Reference


wb = openpyxl.Workbook()
#sheet = wb.get_active_sheet()
sheet=wb.get_sheet_by_name('Sheet')
for i in range(1, 11):  # create some data in column A
    sheet["A" + str(i)] = i


#
refObj = openpyxl.chart.reference.Reference(sheet, min_col=1, min_row=1, max_col=10,
                                            max_row=1)

seriesObj = openpyxl.chart.Series(refObj, title='First series')

chartObj = openpyxl.chart.BarChart()
chartObj.append(seriesObj)
chartObj.drawing = 50 # set position
#chartObj.drawing
#chartObj.drawing. = 300 # set size
#chartObj.drawing.height = 200

sheet.add_chart(chartObj)
wb.save('sampleChart.xlsx')
print('Done.')