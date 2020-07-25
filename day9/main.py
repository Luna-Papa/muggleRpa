from excel_magic.dataset import open_file


file_a = open_file('显示器数据-电商.xlsx')
file_b = open_file('显示器数据-维修部.xlsx')
file_c = open_file("显示器数据-合并.xlsx")

sheet_a = file_a.get_sheet_by_index(0)
sheet_b = file_b.get_sheet_by_index(0)
sheet_c = file_c.get_sheet_by_index(0)

rows_a = sheet_a.get_rows()
rows_b = sheet_b.get_rows()

for row_a, row_b in zip(rows_a, rows_b):
    join_row = row_a + row_b
    filtered_row = join_row.filter_fields(sheet_c.fields)
    sheet_c.append_row(filtered_row)

file_c.save()
