from excel_magic.dataset import open_file, ImageCell
import qrcode

with open_file('客户个人信息.xlsx') as excel_file:
    sheet = excel_file.get_sheet_by_index(0)  # 获取第一个 Sheet
    rows = sheet.get_rows()  # 获取该 Sheet 中的所有行
    for row in rows:  # 遍历每一行
        name = row['姓名'].value  # 把每一行的“姓名”、“电话”、“住址”的值，赋值给 name、phone、address 这 3 个变量
        phone = row['电话'].value
        address = row['住址'].value

        img = qrcode.make(f'{name}\n{phone}\n{address}', border=0, box_size=5)  # 用上面的变量，制造二维码
        img_path = f'{name}.png'  # 预设好二维码图片的储存路径
        img.save(img_path)  # 用该路径保存二维码图片
        row['二维码'] = ImageCell(img_path)  # 读取二维码图片，插入到 Excel 中
