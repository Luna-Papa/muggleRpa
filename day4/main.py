from excel_magic.dataset import open_file
from word_spell import Document

excel_file = open_file("租赁数据.xlsx")
sheet = excel_file.get_sheet_by_index(0)
for row in sheet.get_rows():
    doc = Document("租赁合同模板.docx")
    doc.render_from_template(
        out=f"{row['合同编号']}_{row['承租方姓名']}.docx",
        合同编号=row['合同编号'],
        出租方姓名=row['出租方姓名'],
        出租方身份证=row['出租方身份证'],
        出租方电话=row['出租方电话'],
        承租方姓名=row['承租方姓名'],
        承租方身份证=row['承租方身份证'],
        承租方电话=row['承租方电话'],
        房屋地址=row['房屋地址'],
        租房起始日期=row['租房起始日期'],
        租房终止日期=row['租房终止日期'],
        月租金=row['月租金'],
        保证金=row['保证金'],
    )
