"""Python file used to edit data of a spreadsheet."""

import xlrd

loc = 'resident_information_v2.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

res_list = []
for i in range(sheet.nrows - 1):
    res_list.append(sheet.row_values(i + 1))


