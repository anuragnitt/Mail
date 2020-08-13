import os
import sys
import re
import random
import mimetypes
import base64
import tkinter
from tkinter import filedialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

try :
	import smtplib
except ModuleNotFoundError :
	os.system('pip3 install secure-smtplib')
	import smtplib

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

#######################################################################################################################################

SMTP_HOST = 'smtp.gmail.com'
SMTP_SSL_PORT = 465
SMTP_TLS_PORT = 587

chunk = 1024
header = 16
key_len = 32
iv_len = 16
namelen = 3

app_access_url = 'https://myaccount.google.com/lesssecureapps'

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
		response = input('(Y/N) : ')

		if response not in ['Y', 'y', 'N', 'n'] :
			print('\nTry again.')
			continue

		else :
			break

	return response

#######################################################################################################################################

def send_email(host, port, username, password) :

	os.chdir(os.environ.get('userprofile'))
	sub_dirname = os.path.join(os.environ.get('userprofile'), 'mail-send')
	check_isdir(sub_dirname)
	os.chdir(sub_dirname)

	receiver_list = []
	r_addr = input('\nReceiver\'s E-Mail address : ')
	receiver_list.append(r_addr)

	while True :
		print('\nDo you want to add another receiver ?')
		response = yesno()
		if response in ['Y', 'y'] :
			r_addr = input('\nReceiver\'s E-Mail address : ')
			receiver_list.append(r_addr)
			continue
		else :
			break

	receiver_list = list(dict.fromkeys(receiver_list))

	message = MIMEMultipart()
	message['From'] : username
	message['To'] : ', '.join(receiver_list)

	print('\nDo you want the email to be encrypted ?')
	enc = yesno()
	if enc in ['Y', 'y'] :
		enc = True
		key = get_random_bytes(key_len)
		message['Encryption'] = 'encrypted;encryption=AES-256'
	else :
		enc = False

	print('\nPress ENTER to skip.')
	email_subject = input('SUBJECT : ')
	message['Subject'] = email_subject

	print('\nDo you want an E-Mail body ?')
	response = yesno()

	if response in ['Y', 'y'] :
		f = open('email_body.txt', 'w')
		f.close()
		os.system('notepad email_body.txt')

		while True :
			print('\nDo you want to edit the body ?')
			response = yesno()
			if response in ['Y', 'y'] :
				os.system('notepad email_body.txt')
				continue
			else :
				break

		f = open('email_body.txt', 'r')
		text = f.read()
		f.close()
		os.remove('email_body.txt')

		if enc :
			text = encrypt(text.encode('utf-8'), key)
			text = base64.b64encode(text).decode('utf-8')

		email_body = MIMEText(text, 'plain')
		message.attach(email_body)

	print('\nDo you want to attach files ?')
	response = yesno()

	if response in ['Y', 'y'] :
		win = tkinter.Tk()
		win.geometry('200x50')
		win.title('Non functional window')

		label = tkinter.Label(win, text='IGNORE THIS WINDOW')
		label.pack()

		file_paths = []
		files = filedialog.askopenfilenames(initialdir=os.environ.get('userprofile'), title='Select files to attach')
		file_paths += list(files)

		while True :
			print('\nAdd more files ?')
			response = yesno()

			if response in ['Y', 'y'] :
				files = filedialog.askopenfilenames(initialdir=os.environ.get('userprofile'), title='Select files to attach')
				file_paths += list(files)
				continue

			else :
				break
		win.destroy()
		file_paths = list(dict.fromkeys(file_paths))

		while True :
			print('\nRemove any file (one at a time) ?')
			response = yesno()
			n_files = len(file_paths)

			if response in ['Y', 'y'] :
				print('\nSelected files :\n')
				for i in range(n_files) :
					print('\t{}> {}'.format(i+1, file_paths[i]))
				print('\n')

				while True :
					try :
						x = int(input('File no. to remove : ')) - 1
						if x not in range(n_files) :
							print('\nEnter a valid number.')
							continue

						file_paths.pop(x)
						break

					except ValueError :
						print('\nEnter a valid number.')
						continue
			else :
				break

		print('\nPreparing attachments .....')
		for i in range(len(file_paths)) :
			try :
				print('\n\tProcessing file {} of {}'.format(i+1, len(file_paths)))
				with open(file_paths[i], 'rb') as f :
					file_data = f.read()
					file_name = os.path.basename(f.name)

					if enc :
						print('\tEncrypting .....')
						file_data = encrypt(file_data, key)
						file_name += '.encrypted'
						print('\tDone')

					file_data = MIMEApplication(file_data, Name=file_name)
					file_data['Content-Disposition'] = f'attachment;filename={file_name}'

				message.attach(file_data)

			except :
				print('\nError. Skipping the file.')
		print('\nDone.')

	message = message.as_string()

	print('Connecting to Gmail\'s SMTP server .....')
	try :
		if port == SMTP_TLS_PORT :
			connection = smtplib.SMTP(host, port)
			connection.starttls()
		else :
			connection = smtplib.SMTP_SSL(host, port)
		print('\tConnected.\n')

	except socket.gaierror :
		print('\tCheck your internet connection.\n')
		print('Restart the program to try again.')
		sys.exit()
		dummy_var = input()

	try :
		print('Attempting to identify with the server ....')
		connection.ehlo()
		print('\tIdentification Successful.\n')

	except :
		print('\tIdentification failed.\n\nCheck you internet connection.\n')
		print('Restart the program to try again.')
		sys.exit()
		dummy_var = input()

	try :

		print('Attempting to login .....')
		connection.login(username, password)
		print('\tLogin Successful.\n')

	except :
		print('\tLogin Failed.\n\nAccess to less-secure apps might be "turned off" in your Google account\'s settings.\n\n\tResolve here : {}'.format(app_access_url))
		print('\nRestart the program to try again.')
		sys.exit()
		dummy_var = input()

	if enc :
		print('\nExporting the encryption key .....')
		key_name = ''.join(chr(random.randint(97,122)) for i in range(5))
		with open(f'{key_name}.bin', 'wb') as f :
			f.write(key)
		api_handler = AnonFile('api_key')
		status, file_url = api_handler.upload_file(f'{key_name}.bin')

		if status :
			print('\tExport successful.')
			os.remove(f'{key_name}.bin')
			anon_id = file_url.split('/')[3] + key_name
			print('\n'+'#'*50+f'\nPassword to decrypt the E-Mail : {anon_id}\n'+'#'*50+'\n')
			dummy_var = input('Note it down and press ENTER whenever ready.')
		else :
			print('\fExport failed. Check your internet connection.') # fails only when offline
			export_path = os.path.join(os.environ.env('userprofile'), f'mail-send\\{key_name}.bin')
			os.rename(f'{key_name}.bin', f'{export_path}')
			print(f'\tKey saved to "{export_path}"\n\tYou\'ll have to securely share this file with the respective receivers.')
			dummy_var = input('Press ENTER whenever ready.')

	try :
		print('\nSending E-Mail .....')
		connection.sendmail(username, receiver_list, message)
		print('\nE-Mail Sent.')
		dummy_var = input()

	except :
		print('\nAN UNEXPECTED ERROR OCCURED :/')
		print('\nRestart the program to try again.')
		sys.exit()
		dummy_var = input()

#######################################################################################################################################

def main_function() :

	print('!!! Don\'t be offline (even a millisec) throughout !!!\nEnter your Gmail login credentials.\n')
	username = input('Gmail ID : ')
	password = input('Password : ')

	print('\nWhich type of network traffic encryption to use ?\n')

	while True :
		port = input('(SSL / TLS) : ')
		if port not in ['SSL', 'ssl', 'TLS', 'tls'] :
			print('\nTry again.')
			continue
		else :
			if port in ['SSL', 'ssl'] :
				port = SMTP_SSL_PORT
			else :
				port = SMTP_TLS_PORT
			break

	send_email(SMTP_HOST, port, username, password)

#######################################################################################################################################

main_function()
