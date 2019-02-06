"""Python file used to edit data of a spreadsheet."""

import xlrd

loc = 'resident_information_v2.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print(sheet.cell_value(0, 0))
