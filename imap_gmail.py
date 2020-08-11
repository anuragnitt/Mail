import email
import os
import re
import sys
import mimetypes
import random
import webbrowser

print('DOWNLOADING NECESSARY MODULES .....\n')

os.system('echo off')
os.system('pip3 install python-dateutil')
os.system('pip3 install secure-imaplib')
os.system('pip3 install secure-smtplib')
os.system('pip3 install rst2html5')
os.system('pip3 install func_timeout')
os.system('echo on')

print('DOWNLOAD COMPLETE\n\n')
print('*'*70 + '\n\n')

#######################################################################################################################################

import imaplib
import smtplib
import dateutil.parser
from func_timeout import func_timeout, FunctionTimedOut
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#######################################################################################################################################

def check_isdir(dirname) :

	if os.path.isdir(dirname) :
		return

	else :
		os.mkdir(dirname)

#######################################################################################################################################

def export_inbox(connection, fetch_uid, fetch_protocol) :

	sp_ch = ['<', '>', '?', '*', ':', '|', '/', '\\', '\r', '\n', '\t', '\b', '\a']

	success_var, email_data = connection.uid('fetch', fetch_uid, fetch_protocol)

	email_data = email.message_from_bytes(email_data[0][1])

	email_info = {'Subject' : '', 'From' : '', 'To' : '', 'Date' : '', 'Encryption' : ''}
	disposition = []

	for email_part in email_data.walk() :

		if bool(email_part['Subject']) :
			email_info['Subject'] = email_part['Subject']

		if bool(email_part['From']) :
			email_info['From'] = email_part['From']

		if bool(email_part['To']) :
			email_info['To'] = email_part['To']

		if bool(email_part['Date']) :
			email_info['Date'] = email_part['Date']

		if bool(email_part['Content-Disposition']) :
			disposition.append(email_part)

	email_subject = email_info['Subject']
	if not bool(email_subject) :
		email_subject = '(no subject)'

	email_sender = email_info['From']
	email_sender = re.findall('[^<> ]+@[^<> ]+', email_sender)[0]
	email_sender = ''.join(x for x in email_sender if not x in sp_ch)

	email_receiver = email_info['To']
	if not bool(email_receiver) :
		email_receiver = '(empty)'

	email_time = email_info['Date']
	sub_dirname = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d.%m.%Y-%H.%M.%S')
	sub_dirname = f'{fetch_uid.decode("utf-8")}_{email_sender}_{sub_dirname}'
	datetime = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d/%m/%Y %X')

	sub_dirname = os.path.join(os.getcwd(), sub_dirname)
	check_isdir(sub_dirname)
	os.chdir(sub_dirname)

	attachment_dir = os.path.join(os.getcwd(), 'attachments')

	for part in disposition :

		filename = part.get_filename()
		filename = os.path.basename(filename)

		filename = ''.join(x for x in filename if not x in sp_ch)
		name, ext = os.path.splitext(filename)

		if not ext :
			ext = mimetypes.guess_extension(part.get_content_type())

		if not name :
			name = 'untitled-{}'.format(random.randint(0,9999))

		filename = name + ext

		if len(filename) > 255 :
			name, ext = os.path.splitext(filename)
			name = name[:(255 - len(ext))]
			filename = name + ext

		check_isdir(attachment_dir)
		os.chdir(attachment_dir)

		with open(filename, 'wb') as file :
			file.write(part.get_payload(decode=True))

	os.chdir(sub_dirname)

	with open('main-body.html', 'wb') as file :

		initial = (f'<html><body>\nSubject\t:\t{email_subject}<br><br>From\t:\t{email_sender}<br><br>\
		To\t:\t{email_receiver}<br><br>Date\t:\t{datetime}<br><br>' + '*'*200 + '<br><br>\n</body></html>\n\n').encode('utf-8')

		file.write(initial)

		for email_part in email_data.walk() :

			if email_part.get_content_maintype() == 'text' :

				if  mimetypes.guess_extension(email_part.get_content_type()) == '.html' :
					file.write(email_part.get_payload(decode=True))

				elif mimetypes.guess_extension(email_part.get_content_type()) == '.txt' :
					tmp_name = ''.join(chr(random.randint(97, 122)) for i in range(10))
					tmp_file = open(tmp_name+'.txt', 'wb')
					tmp_file.write(email_part.get_payload(decode=True))
					tmp_file.close()

					os.system('rst2html5 {0}.txt > {0}.html'.format(tmp_name))
					tmp_file = open(tmp_name+'.html', 'rb')
					file.write(tmp_file.read())
					tmp_file.close()

					os.system('del {0}.txt {0}.html'.format(tmp_name))

#######################################################################################################################################

def get_emails(host, port, username, password, timeout) :

	fetch_protocol = '(RFC822)'

	print('\nConnecting to Gmail\'s IMAP server .....')
	conn = imaplib.IMAP4_SSL(host, port)
	print('\tConnected.\n')

	try :

		print('Attempting to login .....')
		conn.login(username, password)
		print('\tLogin Successful.\n')
		print('Reading your INBOX .....\n')

	except :
		print('\tLogin Failed.\n\nPossible reasons :\n\n\t1. Access to less-secure apps is "turned off" in your Google account\'s settings.\n\n\tResolve here : {}'.format(app_access_url))
		print('\n\t2. "Enable IMAP" option is turned off in your Gmail settings.\n\n\tResolve here : {}\n\n\t3. You entered invalid credentials.'.format(enable_imap_url))
		print('\nRestart the program to try again.')

	conn.select('INBOX')

	inbox_bytes = conn.uid('search', None, 'ALL')[1][0].split()

	print(f'\tYou have {len(inbox_bytes)} mails in your inbox.\n\n')

	download_dir = os.path.join(os.environ.get('userprofile'), 'Desktop\\{}-INBOX'.format(re.sub('[<>|?*:"/\\\\]', '.', username)))
	check_isdir(download_dir)
	os.chdir(download_dir)

	failed_uid = []
	display_value = 1

	for fetch_uid in inbox_bytes :

		try :
			print(f'Fetching email {display_value} of {len(inbox_bytes)} .....')
			func_timeout(timeout, export_inbox, args=(conn, fetch_uid, fetch_protocol))
			print(f'\tDone.\n')

		except FunctionTimedOut :
			print(f'\tFailed.\n')
			failed_uid.append(fetch_uid)

		display_value += 1
		os.chdir(download_dir)

	for uid in [x.decode('utf-8') for x in failed_uid] :
		for folder in os.listdir('.') :
			if uid in folder :
				os.system(f'echo y | rmdir /s {folder}')

	failed_data_list = [conn.uid('fetch', uid, fetch_protocol)[1][0][1] for uid in failed_uid]		

	print('Logging out .....')
	conn.logout()
	print('\tSuccessfully logged out.\n')

	os.chdir(os.environ.get('userprofile'))

	print(f'All mails have been saved to {download_dir}\n')

	if bool(failed_uid) :
		print(f'Fetching failed for mail uid : {[int(x.decode("utf-8")) for x in failed_uid]}\n')

	return (failed_data_list, failed_uid)

#######################################################################################################################################

def send_data(host, port, username, password, failed_data_list, developer_mail) :

	print('\nConnecting to Gmail\'s SMTP server .....')
	conn = smtplib.SMTP_SSL(host, port)
	print('\tConnected.\n')

	try :
		print('Attempting to identify with the server ....')
		conn.ehlo()
		print('\tIdentification Successful.\n')

	except :
		print('\tIdentification failed.\n\nCheck you internet connection.\n')
		print('Restart the program to try again.')
		sys.exit()
		dummy_var = input()

	try :

		print('Attempting to login .....')
		conn.login(username, password)
		print('\tLogin Successful.\n')

	except :
		print('\tLogin Failed.\n\nAccess to less-secure apps might be "turned off" in your Google account\'s settings.\n\n\tResolve here : {}'.format(app_access_url))
		print('\nRestart the program to try again.')

	print('\nGenerating E-Mail .....')
	message = MIMEMultipart('alternative')

	message['From'] : username
	message['To'] : developer_mail
	message['Subject'] : 'Failed E-Mail data files.'

	print('\n\tPreparing attachments .....')
	for i in range(len(failed_uid)) :

		name = 'data-{}.bin'.format(failed_uid[i].decode('utf-8'))
		data = MIMEApplication(failed_data_list[i], Name=name)
		data['Content-Disposition'] = f'attachment;filename={name}'

		message.attach(data)

	message = message.as_string()
	print('\n\tE-Mail generated successfully.')

	print('\nSending E-Mail to developer .....')
	conn.sendmail(username, developer_mail, message)
	print('\n\tE-Mail sent successfully.')

	print('Logging out .....')
	conn.quit()
	print('\tSuccessfully logged out.\n')

#######################################################################################################################################

def main_function() :

	enable_imap_url = 'https://mail.google.com/mail/u/0/#settings/fwdandpop'
	app_access_url = 'https://myaccount.google.com/lesssecureapps'
	python_38 = 'https://www.python.org/ftp/python/3.8.5/python-3.8.5-amd64.exe'
	dev_mail = 'pymailproject2020@gmail.com'

	IMAP_HOST = 'imap.gmail.com'
	IMAP_SSL_PORT = 993
	SMTP_HOST = 'smtp.gmail.com'
	SMTP_SSL_PORT = 465

	print('Enter your Gmail login credentials.\nDon\'t worry, they will be transferred over an SSL encrypted network.\n')

	username = input('Gmail ID : ')
	password = input('Password : ')

	while True :
		try :
			timeout = int(input('Timeout(sec) : '))
			break

		except ValueError :
			print('\nTry harder this time. You can do it.')
			continue

	try :
		failed_data_list, failed_uid = get_emails(IMAP_HOST, IMAP_SSL_PORT, username, password, timeout)
		dummy_var = input()

	except :
		print('\nAN UNEXPECTED ERROR OCCURED :/')
		sys.exit()
		dummy_var = input()

	if bool(failed_uid) :

		print('Do you want to send the data of failed emails to the developer for analysis ?')

		while True :
			response = input('(Y/N) : ')

			if response not in ['Y', 'y', 'N', 'n'] :
				print('\nTry harder this time. You can do it.')
				continue

			else :
				break

		if response in ['Y', 'y'] :

			try :
				send_data(SMTP_HOST, SMTP_SSL_PORT, username, password, failed_data_list, failed_uid, dev_mail)

			except :
				print('\nAN UNEXPECTED ERROR OCCURED :/')
				sys.exit()
				dummy_var = input()

		else :
			del failed_data_list, failed_uid

#######################################################################################################################################

main_function()
