"""Python file used to edit data of a spreadsheet."""

import xlrd
import xlsxwriter

loc = 'resident_information_v2.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

res_list = []
for i in range(sheet.nrows - 1):
    res_list.append(sheet.row_values(i + 1))

write_book = xlsxwriter.Workbook('updated_roe_bog.xlsx')
write_sheet = write_book.add_worksheet()
row = 0
col = 0

for name, phone, app, city, county, fema, lat, lon, prop, address in res_list:
    write_sheet.write(row, col, name)
    write_sheet.write(row, col + 1, phone)
    write_sheet.write(row, col + 2, app)
    write_sheet.write(row, col + 3, city)
    write_sheet.write(row, col + 4, fema)
    write_sheet.write(row, col + 5, address)
    row += 1

write_book.close()
