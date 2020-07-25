from excel_magic.dataset import open_file
from word_spell import Document


with open_file('应聘面试信息.xlsx') as excel_file:
    sheet = excel_file.get_sheet_by_index(0)
    rows = sheet.get_rows()
    for row in rows:
        name = row['面试者姓名'].value
        job = row['投递职位'].value
        date = row['面试日期'].value
        time = row['面试时间'].value

        doc = Document('面试通知模板.docx')
        doc.render_from_template(
            面试者姓名=name,
            投递职位=job,
            面试日期=date, 
            面试时间=time, 
            out="./{}.docx".format(name)
        )