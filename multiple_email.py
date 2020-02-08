import smtplib
import ssl
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from auto_certi import make_certificate
from email.mime.image import MIMEImage
import os

subject = ''
sender_email = ""
password = ''

#by using excel
#x1 = pd.ExcelFile('list2.xlsx')
#df = x1.parse('Sheet1')
#files = ['certi.pdf', 'certi.png']


#by using csv
df = pd.read_csv("list2.csv")


for index, row in df.iterrows():
    message = MIMEMultipart()
    message["From"] = sender_email
    message["Subject"] = subject
    print(index, row['Name'], row['email'])
    # body ='Dear '+row['Name'] + '\nCongratulations, your certificate is ready.'

    html = """\
        <html>
        <head>
        <H1>Greetings %s!</H1>
        </head>
        <body>
        <p><h4>Thanks for attending Photoshop Workshop
        <br><br>Kindly find your E-Certificate for the same, attached along with this mail.
        <br>We look forward to seeing you at all our future workshops, seminars and events!
        <br><br>Regards,
        <br>IEEE-VIT Student Branch
        <br><center><br>Connect with us on
        <br><p><a href="http://ieee.vit.edu.in/index.html"><img src="https://png.icons8.com/metro/1600/domain.png" height="50" width="50" hspace="20"><a href="https://www.instagram.com/ieeevit/"><img src="https://images-na.ssl-images-amazon.com/images/I/71VQR1WetdL.png" height="50" width="50" hspace="20"><a href="https://www.facebook.com/IEEEVIT1/"><img src="https://ioufinancial.com/wp-content/uploads/2017/02/facebook.png" height="50" width="50" hspace="20"></a>
        <br><br>
        <font size="2">Ask us anything about programming, meet like minded people, build projects.<br>
        <center>Join the Coders Republic Group now:</font><br><br>
        <a href="https://chat.whatsapp.com/GNjVY5fSZav73fl77vGPj2"><img src="http://www.idroexpert.com/wp-content/uploads/soon-873316_960_720.png" height="50" width="50" hspace="20"></a><br><br>
        <p>For any errors in certificate<a href="https://goo.gl/forms/8dqLFOmrG3KZs65f2"> click here</a></p>
        <img src="cid:myimage" width=800 height=800 />
        <hr>
        </body>
        </html>
        """ % row['Name']

    receiver_email = row['email']
    # print(receiver_email)
    message["To"] = receiver_email
    # print(message["TO"])
    # message.attach(MIMEText(body, "plain"))
    message.attach(MIMEText(html, 'html'))
    make_certificate(row['Name'])
    img_data = open("{}.png".format(row['Name']), 'rb').read()
    img = MIMEImage(img_data, 'jpg')
    img.add_header('Content-Id', '<myimage>')
    img.add_header("Content-Disposition", "inline", filename="{}.png".format(row['Name']))
    message.attach(img)
    '''
    for filename in files:
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",f"attachment; filename= {filename}",)
        message.attach(part)
        '''
    with open("{}.png".format(row['Name']), "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {'{}.png'.format(row['Name'])}", )
    message.attach(part)
    text = message.as_string()
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
        print('email send to ', receiver_email)
    del message
    os.remove("{}.png".format(row['Name']))

