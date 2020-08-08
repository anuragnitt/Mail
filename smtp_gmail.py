from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import os

host = 'smtp.gmail.com'
SMTP_SSL_PORT = 465
SMTP_TLS_PORT = 587

#server = smtplib.SMTP(host, SMTP_TLS_PORT)
server = smtplib.SMTP_SSL(host, SMTP_SSL_PORT)

def send_email(from_mail, passwd, to_mail_list, subject, message_text, message_text_type, attachment_list) :

	for mail in to_mail_list :

		message = MIMEMultipart('alternative')
		message['Subject'] = subject
		message['From'] = from_mail
		message['To'] = mail
		message.attach(MIMEText(message_text, message_text_type))

		for file in attachment_list :

			f = open(file, 'rb')
			file_data = f.read()
			name = os.path.basename(f.name)
			f.close()

			file_data = MIMEApplication(file_data, Name=name)
			file_data['Content-Disposition'] = f'attachment; filename={name}'
			message.attach(file_data)

		#server.connect(host, SMTP_TLS_PORT)
		server.connect(host, SMTP_SSL_PORT)
		server.ehlo()
		#server.starttls()
		server.login(from_mail, passwd)
		server.sendmail(from_mail, mail, message.as_string())
		server.quit()

uname = 'yourmail@gmail.com'
passwd = 'your_passwd'

to_mail = ['''all receivers' list''']

html_text = """\
		<html>
		   <head></head>
		      <body>
		         <p><b>Hey there!</b><br><br>
		            Its me, John Doe. This email was sent by an automated python script I wrote.<br><br>
		            If you like this work, be sure to follow me on <a href=https://www.instagram.com/username>Instagram</a>.
		            <br><br>Random files have been attached.
		         </p>
		      </body>
		</html>"""

file_list = ['''list of all file locations''']

send_email(uname, passwd, to_mail, 'Automated Scipted Mail', html_text, 'html', file_list)

del server
