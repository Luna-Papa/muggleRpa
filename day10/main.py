from excel_magic.dataset import open_file


file = open_file('fans.xlsx')
sheet = file.get_sheet_by_index(0)
rows = sheet.get_rows()
fans_num = len(rows)
group_size = 30
num = 0

for i in range(0, fans_num, group_size):
    num += 1
    group_rows = rows[i:i+group_size]
    new_file = open_file(f"result/「初音ミク 粉丝群」群{num}.xlsx")
    new_sheet = new_file.add_sheet("sheet1", sheet.fields)
    new_sheet.append_rows(group_rows)
    new_file.save(backup=False)
