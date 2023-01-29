from email.mime.text import MIMEText
from email.generator import Generator
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
# create a simple message
msg = MIMEMultipart()

content = '''This is a simple message.
And a very simple one. 
我是中文字 拉拉拉'''
msg['Subject'] = 'Simple message'
msg['From'] = 'sender@sending.domain'
msg['To'] = 'rcpt@receiver.domain'
# 寫入信件內容
txt = MIMEText(content, 'plain', 'utf-8')
msg.attach(txt)
# 寫入檔案內容
out_file = MIMEApplication(open('test.txt', "rb").read())
out_file.add_header('Content-Disposition',
                    'attachment', filename="test.text")
msg.attach(out_file)

# open a file and save mail to it
with open('filename111112.eml', 'w') as out:
    
    gen = Generator(out)
    gen.flatten(msg)


