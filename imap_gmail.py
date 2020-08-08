import imaplib
import email
import os
import re
import dateutil.parser
import sys
import webbrowser

enable_imap_url = 'https://mail.google.com/mail/u/0/#settings/fwdandpop'
app_access_url = 'https://myaccount.google.com/lesssecureapps'

webbrowser.open(app_access_url, new=1)
webbrowser.open(enable_imap_url, new=1)

print('\t\tDOWNLOADING NECESSARY MODULES\n\n')

os.system('pip3 install python-dateutil')
os.system('pip3 install secure-imaplib')

print('\n\t\tDOWNLOAD COMPLETE\n\n')
print('*'*70 + '\n\n')

def check_isdir(dirname) :

	if os.path.isdir(dirname) :
		return

	else :
		os.mkdir(dirname)

def export_inbox(host, port, username, password, fetch_protocol) :

	print('\nConnecting to Gmail server .....')
	conn = imaplib.IMAP4_SSL(host, port)
	print('\tConnected.\n')

	try :

		print('Attempting to login .....')
		conn.login(username, password)
		print('\tLogin Successful.\n')

	except :
		print('\tLogin Failed.\n\nPossible reasons :\n\n\t1. Access to less-secure apps is "turned off" in your Google account.\n\n\tResolve here : {}'.format(app_access_url))
		print('\n\t2. "Enable IMAP" option is turned off in your Gmail settings.\n\n\tResolve here : {}\n\n\t3. You entered invalid credentials.'.format(enable_imap_url))
		print('\nRestart the program to try again.')

		dummy_var = input()
		sys.exit()

	download_dir = os.path.join(os.environ.get('userprofile'), 'Desktop\\{}-INBOX'.format(username))

	check_isdir(download_dir)
	os.chdir(download_dir)

	conn.select('INBOX')

	inbox_bytes = conn.uid('search', None, 'ALL')[1][0].split()

	email_counter = len(inbox_bytes)
	display_value = 1

	print(f'You have {email_counter} mails in your inbox.\n\n')

	for fetch_uid in inbox_bytes :

		print(f'Fetching mail {display_value} of {len(inbox_bytes)} .....')

		success_var, email_data = conn.uid('fetch', fetch_uid, fetch_protocol)

		email_data = email.message_from_bytes(email_data[0][1])

		email_info = []

		for email_part in email_data.walk() :

			if bool(email_part['Subject']) :
				email_info.append(email_part['Subject'])

			if bool(email_part['From']) :
				email_info.append(email_part['From'])

			if bool(email_part['To']) :
				email_info.append(email_part['To'])

			if bool(email_part['Date']) :
				email_info.append(email_part['Date'])

			if bool(email_part.get_filename()) :
				email_info.append(email_part)

		email_subject = email_info[0]

		email_sender = email_info[1]
		email_sender = re.findall('[^<> ]+@[^<> ]+', email_sender)[0]

		email_receiver = email_info[2]

		email_time = email_info[3]
		sub_dirname = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d-%m-%Y_%H-%M-%S')
		sub_dirname = f'{email_counter}_{email_sender}_{sub_dirname}'
		datetime = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d/%m/%Y %X')

		sub_dirname = os.path.join(os.getcwd(), sub_dirname)
		check_isdir(sub_dirname)
		os.chdir(sub_dirname)

		attachment_count = len(email_info) - 4

		if attachment_count :

			attachment_dir = os.path.join(os.getcwd(), 'attachments')
			check_isdir(attachment_dir)
			os.chdir(attachment_dir)

		for i in range(attachment_count) :

			that_part = email_info[4 + i]

			filename = os.path.basename(that_part.get_filename())

			with open(filename, 'wb') as file :
				file.write(that_part.get_payload(decode=True))

		os.chdir(sub_dirname)

		with open('main-body.html', 'wb') as file :

			file.write('<html><body>\n'.encode('utf-8'))
			file.write(f'Subject\t:\t{email_subject}<br><br>'.encode('utf-8'))
			file.write(f'From\t:\t{email_sender}<br><br>'.encode('utf-8'))
			file.write(f'To\t:\t{email_receiver}<br><br>'.encode('utf-8'))
			file.write(f'Date\t:\t{datetime}<br><br>'.encode('utf-8'))
			file.write(('*'*200 + '<br><br>\n').encode('utf-8'))
			file.write('</body></html>\n\n'.encode('utf-8'))

			for email_part in email_data.walk() :

				if email_part.get_content_maintype() == 'text' :
					file.write(email_part.get_payload(decode=True))

		print('\tDone.\n')

		os.chdir(download_dir)

		email_counter -= 1
		display_value += 1

	print('Logging out .....')
	conn.logout()
	print('\tSuccessfully logged out.\n')

	os.chdir(os.environ.get('userprofile'))

	print(f'All mails have been saved to {download_dir}')

IMAP_HOST = 'imap.gmail.com'
IMAP_SSL_PORT = 993

#username = os.environ.get('agmail_addr')
#password = os.environ.get('agmail_pass')

print('Enter your Gmail login credentials.\nDon\'t worry, they will be transferred over an SSL encrypted network.\n')

username = input('Gmail ID : ')
password = input('Password : ')

fetch_protocol = '(RFC822)'

try :

	export_inbox(IMAP_HOST, IMAP_SSL_PORT, username, password, fetch_protocol)
	dummy_var = input()

except :

	dummy_var = input()
	sys.exit()