import os
from excel_magic.dataset import open_file


file_list = os.listdir()
excel_file = open_file('./招福银行客户信息总表.xlsx')
for f in file_list:
    if ".xlsx" in f:
        excel_file.merge_file(f)

excel_file.save()

