import os
import sys
import re
import random
import mimetypes
import base64
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tkinter
from tkinter import filedialog

try :
	import imaplib
except ModuleNotFoundError :
	os.system('pip3 install secure-imaplib')
	import imaplib

try :
	import smtplib
except ModuleNotFoundError :
	os.system('pip3 install secure-smtplib')
	import smtplib

try :
	import dateutil.parser
except ModuleNotFoundError :
	os.system('pip3 install python-dateutil')
	import dateutil.parser

try :
	from func_timeout import func_timeout, FunctionTimedOut
except ModuleNotFoundError :
	os.system('pip3 install func_timeout')
	from func_timeout import func_timeout, FunctionTimedOut

try :
	from Crypto.Cipher import AES
	from Crypto.Random import get_random_bytes
except ModuleNotFoundError :
	os.system('pip3 install pycryptodome')
	from Crypto.Cipher import AES
	from Crypto.Random import get_random_bytes

try :
	from anonfile.anonfile import AnonFile
except ModuleNotFoundError :
	os.system('pip3 install anonfile')
	from anonfile.anonfile import AnonFile

try :
	import rst2html5_
except ModuleNotFoundError :
	os.system('pip3 install rst2html5')
	import rst2html5_

#######################################################################################################################################

enable_imap_url = 'https://mail.google.com/mail/u/0/#settings/fwdandpop'
app_access_url = 'https://myaccount.google.com/lesssecureapps'

chunk = 1024
header = 16
key_len = 32
iv_len = 16
namelen = 3

#######################################################################################################################################

def encrypt(byte_data, key) :

	iv = get_random_bytes(iv_len)
	org_sz = len(byte_data)
	pad_sz = chunk - org_sz%chunk

	if org_sz%chunk != 0 :
		byte_data += (chr(0).encode('utf-8'))*pad_sz

	n_blocks = len(byte_data)//chunk

	engine = AES.new(key, AES.MODE_CBC, iv)

	eb = f'{org_sz:<{header}}'.encode('utf-8')
	eb += iv
	for i in range(n_blocks) :

		block = byte_data[:chunk]
		byte_data = byte_data[chunk:]
		block = engine.encrypt(block)
		eb += block

	return eb

#######################################################################################################################################

def decrypt(byte_data, key) :

	org_sz = int(byte_data[:header].decode('utf-8'))
	byte_data = byte_data[header:]

	iv = byte_data[:iv_len]
	byte_data = byte_data[iv_len:]

	n_blocks = len(byte_data)//chunk

	engine = AES.new(key, AES.MODE_CBC, iv)

	db = b''
	for i in range(n_blocks) :

		block = byte_data[:chunk]
		byte_data = byte_data[chunk:]
		block = engine.decrypt(block)
		db += block

	pad_sz = chunk - org_sz%chunk
	db = db[::-1][pad_sz:][::-1]

	return db

#######################################################################################################################################

def check_isdir(dirname) :

	if os.path.isdir(dirname) :
		return

	else :
		os.mkdir(dirname)

#######################################################################################################################################

def yesno() :

	while True :
		response = input('\t(Y/N) : ')

		if response not in ['Y', 'y', 'N', 'n'] :
			print('\n\tTry again.')
			continue

		else :
			break

	return response

#######################################################################################################################################

def export_inbox(connection, fetch_uid, fetch_protocol) :

	sp_ch = ['<', '>', '?', '*', ':', '|', '/', '"', '\\', '\r', '\n', '\t', '\b', '\a']

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

		if bool(email_part['Encryption']) :
			email_info['Encryption'] = email_part['Encryption']

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

	key = None
	if bool(email_info['Encryption']) :

		print('\n\tThis E-Mail is secured by encryption.')
		print(f'\n\tHINTS:\n\t\tSender : {email_sender}\n\t\tSubject : {email_subject}')
		print('\n\tYou have only 5 attempts to enter the correct password.')
		print('\tAfter that the E-Mail will be deleted.')
		print('\n\tBrowse for a key containing file instead (5 attempts) ?')
		response = yesno()

		if response in ['Y', 'y'] :

			win = tkinter.Tk()
			win.geometry('200x50')
			win.title('Non functional window')

			label = tkinter.Label(win, text='IGNORE THIS WINDOW')
			label.pack()

			for i in range(5) :
				print(f'\n\t[Attempt {i+1} of 5]')
				key_path = filedialog.askopenfilename(initialdir=os.environ.get('userprofile'), title='Select the file having the key', filetypes=[('Binary', '*.bin')])
				with open(key_path, 'rb') as f :
					key = f.read()
				#os.remove(key_path)
				if len(key) == 32 :
					print('\tKEY ACCEPTED\n\n\tDecrypting .....')
					break
				else :
					print('\tWRONG KEY')
					key = None
					continue
			win.destroy()

		else :

			for i in range(5) :
				print(f'\n\t[Attempt {i+1} of 5]')
				anon_id = input('\n\tPASSWORD : ')
				if len(anon_id) == 15 :
					a = 1
					for x in anon_id : a = a*x.isalnum()
					if a :
						key_name = anon_id[-5:]
						anon_id = anon_id[:-5]
						#api_res = 'temp-{}'.format(random.randint(0,9999))
						api_handler = AnonFile('api_key')
						sys.stdout = open(os.devnull, 'w')
						api_handler.download_file(f'https://anonfiles.com/{anon_id}/{key_name}_bin')
						sys.stdout = sys.__stdout__
						#with open(api_res, 'r') as f :
						#	dl_res = f.read()
						#os.remove(api_res)

						#if 'Error -- 403: Forbidden' in dl_res :
						#	print('\tWRONG PASSWORD')
						#	continue
						if os.path.exists(f'{key_name}.bin') :
							with open(f'{key_name}.bin', 'rb') as f :
								key = f.read()
							os.remove(f'{key_name}.bin')
							print('\tPASSWORD ACCEPTED\n\n\tDecrypting .....')
							break
						else :
							print('\tWRONG PASSWORD')
							continue
					else :
						print('\tWRONG PASSWORD')
						continue
				else :
					print('\tWRONG PASSWORD')
					continue

		if not key :
			print('\n\tAll attempts failed.\n\tDeleting the E-Mail .....')
			connection.uid('store', fetch_uid, '+FLAGS', '\\Deleted')
			connection.expunge()
			print('\tE-Mail deleted.')
			return

	# existence of key implies existence of encryption and vice-versa otherwise the function would have quit by now.

	sub_dirname = os.path.join(os.getcwd(), sub_dirname)
	check_isdir(sub_dirname)
	os.chdir(sub_dirname)

	attachment_dir = os.path.join(os.getcwd(), 'attachments')

	for part in disposition :

		filename = part.get_filename()

		if filename == None :
			filename = ''
		else :
			filename = os.path.basename(filename)
			if key :
				if filename.endswith('.encrypted') :
					filename = filename.replace('.encrypted', '')
				else :
					print('\nSomeone messed with the mail-sending source.\nAttachment names may be affected.')

		filename = ''.join(x for x in filename if not x in sp_ch)
		name, ext = os.path.splitext(filename)

		if not ext :
			ext = mimetypes.guess_extension(part.get_content_type())

		if not name :
			if 'attachment' in part['Content-Disposition'] :
				name = 'untitled-attachment-{}'.format(random.randint(0,9999))

			elif 'inline' in part['Content-Disposition'] :
				name = 'untitled-inline-{}'.format(random.randint(0,999))

		filename = name + ext

		if len(filename) > 255 :
			name, ext = os.path.splitext(filename)
			name = name[:(255 - len(ext))]
			filename = name + ext

		check_isdir(attachment_dir)
		os.chdir(attachment_dir)

		with open(filename, 'wb') as file :
			payload = part.get_payload(decode=True)
			if key :
				payload = decrypt(payload, key)
			file.write(payload)

	os.chdir(sub_dirname)

	with open('main-body.html', 'wb') as file :

		initial = (f'<html><body>\nSubject\t:\t{email_subject}<br><br>From\t:\t{email_sender}<br><br>\
		To\t:\t{email_receiver}<br><br>Date\t:\t{datetime}<br><br>' + '*'*200 + '<br><br>\n</body></html>\n\n').encode('utf-8')

		file.write(initial)

		for email_part in email_data.walk() :

			if email_part.get_content_maintype() == 'text' :

				if  mimetypes.guess_extension(email_part.get_content_type()) == '.html' :
					payload = email_part.get_payload(decode=True)
					if key :
						payload = base64.b64decode(payload)
						payload = decrypt(payload, key)
					file.write(payload)

				elif mimetypes.guess_extension(email_part.get_content_type()) == '.txt' :
					tmp_name = ''.join(chr(random.randint(97, 122)) for i in range(10))
					tmp_file = open(f'{tmp_name}.txt', 'wb')
					payload = email_part.get_payload(decode=True)
					if key :
						payload = base64.b64decode(payload)
						payload = decrypt(payload, key)
					tmp_file.write(payload)
					tmp_file.close()

					os.system('rst2html5 {0}.txt > {0}.html'.format(tmp_name))
					tmp_file = open(f'{tmp_name}.html', 'rb')
					file.write(tmp_file.read())
					tmp_file.close()

					os.remove(f'{tmp_name}.txt')
					os.remove(f'{tmp_name}.html')

#######################################################################################################################################

def get_emails(host, port, username, password, timeout) :

	fetch_protocol = '(RFC822)'

	print('\nConnecting to Gmail\'s IMAP server .....')
	try :
		conn = imaplib.IMAP4_SSL(host, port)
		print('\tConnected.\n')

	except socket.gaierror :
		print('\tCheck your internet connection.\n')
		print('Restart the program to try again.')
		sys.exit()
		dummy_var = input()

	try :

		print('Attempting to login .....')
		conn.login(username, password)
		print('\tLogin Successful.\n')
		print('Reading your INBOX .....\n')

	except :
		print('\tLogin Failed.\n\nPossible reasons :\n\n\t1. Access to less-secure apps is "turned off" in your Google account\'s settings.\n\n\tResolve here : {}'.format(app_access_url))
		print('\n\t2. "Enable IMAP" option is turned off in your Gmail settings.\n\n\tResolve here : {}\n\n\t3. You entered invalid credentials.'.format(enable_imap_url))
		print('\nRestart the program to try again.')
		sys.exit()
		dummy_var = input()

	conn.select('INBOX')

	inbox_bytes = conn.uid('search', None, 'ALL')[1][0].split()[::-1]

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
				os.system(f'echo y | rmdir /s {folder}') # Windows
				#os.system(f'sudo rm -r {folder}') # Linux

	failed_data_list = [conn.uid('fetch', uid, fetch_protocol)[1][0][1] for uid in failed_uid]		

	print('Logging out .....')
	conn.logout()
	print('\tSuccessfully logged out.\n')

	os.chdir(os.environ.get('userprofile'))

	print(f'All mails have been saved to {download_dir}\n')

	if bool(failed_uid) :
		print(f'Fetch failure for these mail uid(s) : {[int(x.decode("utf-8")) for x in failed_uid]}\n')

	return (failed_data_list, failed_uid)

#######################################################################################################################################

def send_data(host, port, username, password, failed_data_list, developer_mail) :

	print('\nConnecting to Gmail\'s SMTP server .....')
	try :
		conn = smtplib.SMTP_SSL(host, port)
		conn.starttls()
		print('\tConnected.\n')

	except socket.gaierror :
		print('\tCheck your internet connection.\n')
		print('Restart the program to try again.')
		sys.exit()
		dummy_var = input()

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
		sys.exit()
		dummy_var = input()

	print('\nGenerating E-Mail .....')
	message = MIMEMultipart('alternative')

	message['From'] : username
	message['To'] : developer_mail
	message['Subject'] : 'Failed E-Mail data files.'
	message['Encryption'] : 'encrypted;encryption=AES-256'
	key = get_random_bytes(32)

	print('\n\tPreparing encrypted attachments .....')
	for i in range(len(failed_uid)) :

		name = 'data-{}.bin.encrypted'.format(failed_uid[i].decode('utf-8'))
		data = encrypt(failed_data_list[i], key)
		data = MIMEApplication(data, Name=name)
		data['Content-Disposition'] = f'attachment;filename={name}'
		message.attach(data)

	print('\nExporting the encryption key .....')
	key_name = ''.join(chr(random.randint(97,122)) for i in range(5))
	with open(f'{key_name}.bin', 'wb') as f :
		f.write(key)
	api_handler = AnonFile('api_key')
	status, file_url = api_handler.upload_file(f'{key_name}.bin')

	if status :
		print('\tExport successful.')
		os.remove(f'{key_name}.bin')
		anon_id = (file_url.split('/')[3] + key_name).encode('utf-8')
		anon_id = base64.b64encode(anon_id).decode('utf-8')
		anon_id = MIMEText(anon_id, 'plain')
		message.attach(anon_id)
	else :
		print('\tExport failed. Check your internet connection.') # fails only when offline.
		export_path = os.path.join(os.environ.env('userprofile'), f'mail-send\\{key_name}.bin')
		os.rename(f'{key_name}.bin', f'{export_path}')
		print(f'\tKey saved to "{export_path}"\n\tMail this file to {developer_mail}')
		dummy_var = input('Press ENTER whenever ready.')


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

	dev_mail = 'pymailproject2020@gmail.com'

	IMAP_HOST = 'imap.gmail.com'
	IMAP_SSL_PORT = 993
	SMTP_HOST = 'smtp.gmail.com'
	SMTP_SSL_PORT = 465
	SMTP_TLS_PORT = 587

	print('!!! Don\'t be offline (even a millisec) throughout !!!\nEnter your Gmail login credentials.\n')

	username = input('Gmail ID : ')
	password = input('Password : ')
	#timeout = 300
	print('\nMaximum time to wait before skipping an E-Mail (Recommended - 300sec) ?')
	while True :
		try :
			timeout = int(input('Timeout(sec) : '))
			break

		except ValueError :
			print('\nTry again.')
			continue

	try :
		failed_data_list, failed_uid = get_emails(IMAP_HOST, IMAP_SSL_PORT, username, password, timeout)
		dummy_var = input()

	except :
		print('\nAN UNEXPECTED ERROR OCCURED :/')
		sys.exit()
		dummy_var = input()

	if bool(failed_uid) :

		print('\nDo you want to send the data of failed emails to the developer for analysis ?')
		print('(It will be encrypted while sharing and will be deleted immediately after analysis)')
		response = yesno()

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
