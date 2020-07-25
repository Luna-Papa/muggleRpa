from mail_mojo.QQMail import QQMail
from mail_mojo.message import Message
from .config import ADDRESS, TOKEN


auto_client = QQMail(address=ADDRESS, access_token=TOKEN)
auto_client.init_imap()
auto_client.init_wait()

msg_text = '''
您好：

    库存信息表已发送给您。
    
    您可以扫描附件二维码查看详情。

祝好!
'''

while True:
    print('服务开始，接收邮件中...')
    new_mail = auto_client.wait_for_new_mail()
    print('>>> 新邮件！')
    print(f'>> 收到来自 {new_mail.FROM} 的邮件，主题为：{new_mail.SUBJECT}')
    if '库存' in new_mail.CONTENT:
        msg = Message(from_name=ADDRESS, to_name=new_mail.FROM, subject=f'自动回复: {new_mail.SUBJECT}')
        msg.add_text(msg_text)
        msg.add_image('二维码.png')
        msg.add_attachment('库存信息.xlsx')

        auto_client.send(new_mail.FROM, msg)
        print(f'> 已回复邮件给 {new_mail.FROM} ✔️')
    else:
        print(f'> 无关邮件，未回复 x')
