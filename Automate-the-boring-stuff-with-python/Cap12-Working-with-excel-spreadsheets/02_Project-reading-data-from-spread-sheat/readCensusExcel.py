#!/usr/bin/env python3

"""readCensusEcel.py - tabulates population and number  of census tracts for each county."""

import openpyxl, pprint

print("Opening workbook...")
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb.get_sheet_by_name('Population by Census Tract')
countyData={}

# TODO: Fill in countyData with each county's population and tracts.
print("Reading rows...")


# for row in range(2, sheet.get_highest_row() + 1):
for row in range(2, sheet.max_row + 1):
    # Each row in the spreadsheet has data for one census tract.
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value

    countyData.setdefault(state, {})
    countyData[state].setdefault(county, {'tracts': 0, 'pop':0})
    countyData[state][county]['tracts'] += 1
    countyData[state][county]['pop'] += int(pop)

print('Writing results...')
resultFile = open("census2010.py", "w")
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print('Done!')

# TODO: Open a new text file  and write the content of countyData to it.
