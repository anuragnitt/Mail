from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import messagebox
import smtplib
import sys

host = 'smtp.gmail.com'
SMTP_SSL_PORT = 465
SMTP_TLS_PORT = 587

username = 'pymailproject2020@gmail.com'
passwd = 'pymail2020'

from_mail = username
to_mail = ['anuraggoyal.aawwrr@gmail.com', 'shanu94301@gmail.com', 'agarwalkartikey21@gmail.com', 'meet1655@gmail.com']

conn = smtplib.SMTP_SSL(host, SMTP_SSL_PORT)
conn.ehlo()

message = MIMEMultipart('alternative')
message['Subject'] = 'Automated Script'
message['From'] = from_mail

conn.login(username, passwd)

for mail in to_mail :
	
	message['To'] = mail

	html_text = """\
	<html>
	   <head></head>
	      <body>
	         <p><b>Hey there!</b><br><br>
	            Its me, Anurag. This email was sent by an automated python script I wrote.<br><br>
	            If you like this work, be sure to follow me on <a href=https://www.instagram.com/anurag_goyal.awr>Instagram</a>.
	         </p>
	      </body>
	   </html>"""

	html_text = MIMEText(html_text, 'html')

	message.attach(html_text)
	
	conn.sendmail(from_mail, mail, message.as_string())

conn.quit()