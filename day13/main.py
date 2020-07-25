import time

from excel_magic.dataset import open_file
from helium import *  # 导入 helium 库中的全部函数

path = 'file:///D:/project/muggleRpa/day13/库存提交系统.html'  # 本地 HTML 文件的路径，可以用浏览器打开 HTML，复制地址栏即可
start_chrome(path)
all_text = find_all(Text())
id_list = [i.web_element.text for i in all_text if i.web_element.text.isnumeric()]

sh = open_file("库存信息.xlsx").get_sheet_by_index(0)
result = sh.filter(lambda row: str(row["商品编号"]) in id_list)

for row in result:
    order_id = str(row["商品编号"].value)
    ipt = TextField(to_right_of=order_id)
    write(row['存量'].value, into=ipt)
    cb = CheckBox(to_right_of=ipt)
    if int(row['存量'].value) < 100:
        click(cb)

time.sleep(1)
click(Button("提交"))
time.sleep(2)
kill_browser()
