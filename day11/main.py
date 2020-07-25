from excel_magic.dataset import open_file
from urllib.request import urlretrieve


excel = open_file('图片库.xlsx')
sheet = excel.get_sheet_by_index(0)
rows = sheet.get_rows()

for row in rows:
    name = row['人物名称'].value
    url = row['下载链接'].value
    path = f'result/{name}.jpeg'
    urlretrieve(url, path)
    print(f'{name} 已下载 ✔️')
