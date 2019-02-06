"""Python file used to edit data of a spreadsheet."""

import xlrd
import xlsxwriter

loc = 'resident_information_v2.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

res_list = []
for i in range(sheet.nrows - 1):
    row_val = sheet.row_values(i + 1)
    address = row_val.pop(9)
    add_list = address.split(' ')
    house_num = add_list[0]
    street_add_list = add_list[1:]
    street_add = ' '.join(street_add_list)
    row_val.append(house_num)
    row_val.append(street_add)
    res_list.append(row_val)

write_book = xlsxwriter.Workbook('updated_roe_bog.xlsx')
write_sheet = write_book.add_worksheet()
row = 0
col = 0

headers = ['Applicant Name', 'Phone Number', 'Application Number', 'City',
           'FEMA ID', 'House Number', 'Street Address']

write_sheet.write(0, 0, headers[0])
write_sheet.write(0, 1, headers[1])
write_sheet.write(0, 2, headers[2])
write_sheet.write(0, 3, headers[3])
write_sheet.write(0, 4, headers[4])
write_sheet.write(0, 5, headers[5])
write_sheet.write(0, 6, headers[6])

row += 1

for name, phone, app, city, county, fema, lat, lon, prop, house, street in res_list:
    write_sheet.write(row, col, name)
    write_sheet.write(row, col + 1, phone)
    write_sheet.write(row, col + 2, app)
    write_sheet.write(row, col + 3, city)
    write_sheet.write(row, col + 4, fema)
    write_sheet.write(row, col + 5, house)
    write_sheet.write(row, col + 6, street)
    row += 1

write_book.close()
