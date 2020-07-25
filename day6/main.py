from mail_mojo.QQMail import QQMail
from mail_mojo.message import Message
from excel_magic.dataset import open_file
from .config import ADDRESS, TOKEN


subject = '【端午节放假通知】'
mail_text = '''
亲爱的{}同事，您好：

　　根据《国务院办公厅公布2020年放假安排》，结合公司的实际情况，端午节期间放假安排如下：

　　2020年端午节放假，2020年06月25日至06月27日放假三天。

　　注意事项：

　　1、请各部门负责人提前组织好放假前安全检查，并做好防火、防盗排查;

　　2、放假期间需要外出的员工请注意安全防护，避免意外事故发生。

　　3、公司今年的端午礼品是六芳斋粽子礼盒，每人一盒。

　　祝您节日愉快，身体健康!
'''

client = QQMail(address=ADDRESS, access_token=TOKEN)

excel_file = open_file('员工表.xlsx')
sheet = excel_file.get_sheet_by_index(0)
for row in sheet.get_rows():
    name = row['姓名']
    mail = row['邮箱']

    msg = Message(from_name=ADDRESS, to_name=mail.value, subject=subject)
    msg.add_text(mail_text.format(name))
    msg.add_attachment('端午放假时间.png')
    client.send(to_address=mail.value, msg=msg)
    print(f'{name} {mail} ✔️')
