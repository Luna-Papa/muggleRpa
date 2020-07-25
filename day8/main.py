from excel_magic.dataset import open_file,Style
from excel_magic.diff import diff


total_file = open_file('库存信息总表.xlsx')
local_file = open_file('库存信息子表.xlsx')
check_file = local_file.duplicate("错误检查.xlsx")

sheet_a = local_file.get_sheet_by_index(0)
sheet_b = total_file.get_sheet_by_index(0)
sheet_c = check_file.get_sheet_by_index(0)

diff_result = diff(sheet_a, sheet_b)
diff_sheet = diff_result.not_found_in_b

diff_rows = diff_result.not_found_in_b.get_rows()   # 找到有问题的 row
highlight_style = Style(fill_color='#FFC7CD', font_color='#C65F60')
sheet_c.highlight(diff_rows, highlight_style)   # 在 check_file 中给这些问题 row 标红
check_file.sheets.append(diff_sheet)   # 同时为问题 row 新建一个 sheet ，添加到 check_file 中

check_file.save()
