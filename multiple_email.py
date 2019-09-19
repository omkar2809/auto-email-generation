import smtplib
import ssl
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from auto_certi import make_certificate

subject = ""
sender_email = ""
password = ''

x1 = pd.ExcelFile('list2.xlsx')
df = x1.parse('Sheet1')
files = ['certi.pdf', 'certi.png']

for index, row in df.iterrows():
    message = MIMEMultipart()
    message["From"] = sender_email
    message["Subject"] = subject
    print(row['Name'],row['email'])
    body ='Dear '+row['Name'] + '\nCongratulations, your certificate is ready.'

    html = """\
    <html>
        
    </html>
    """

    receiver_email = row['email']
    print(receiver_email,'\n')
    message["To"] = receiver_email
    print(message["TO"])
    message.attach(MIMEText(body, "plain"))
    message.attach(MIMEText(html, 'html'))
    make_certificate(row['Name'])

    for filename in files:
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",f"attachment; filename= {filename}",)
        message.attach(part)

    text = message.as_string()
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
        print('email send to ', receiver_email)

    del message

