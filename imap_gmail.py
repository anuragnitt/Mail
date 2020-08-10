import imaplib
import email
import os
import re
import dateutil.parser
import sys

#######################################################################################################################################

enable_imap_url = 'https://mail.google.com/mail/u/0/#settings/fwdandpop'
app_access_url = 'https://myaccount.google.com/lesssecureapps'

print('\t\tDOWNLOADING NECESSARY MODULES\n\n')

os.system('pip3 install python-dateutil')
os.system('pip3 install secure-imaplib')

print('\n\t\tDOWNLOAD COMPLETE\n\n')
print('*'*70 + '\n\n')

#######################################################################################################################################

def check_isdir(dirname) :

	if os.path.isdir(dirname) :
		return

	else :
		os.mkdir(dirname)

#######################################################################################################################################

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

	download_dir = os.path.join(os.environ.get('userprofile'), 'Desktop\\{}-INBOX'.format(re.sub('[<>|?*:"/\\\\]', '.', username)))

	check_isdir(download_dir)
	os.chdir(download_dir)

	conn.select('INBOX')

	inbox_bytes = conn.uid('search', None, 'ALL')[1][0].split()[::-1]

	display_value = 1

	print(f'You have {len(inbox_bytes)} mails in your inbox.\n\n')

	for fetch_uid in inbox_bytes :

		print(f'Fetching mail {display_value} of {len(inbox_bytes)} .....')

		success_var, email_data = conn.uid('fetch', fetch_uid, fetch_protocol)

		email_data = email.message_from_bytes(email_data[0][1])

		email_info = {'Subject' : '', 'From' : '', 'To' : '', 'Date' : ''}
		disposition = []

		for email_part in email_data.walk() :

			if type(email_part['Subject']) == type('') : # email without subject will have an empty string as subject.

				if len(email_part['Subject']) > 0 :
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
		email_sender = re.sub('[<>|?*:"/\\\\]', '.', email_sender)

		email_receiver = email_info['To']
		if not bool(email_receiver) :
			email_receiver = '(empty)'

		email_time = email_info['Date']
		sub_dirname = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d.%m.%Y-%H.%M.%S')
		sub_dirname = f'{display_value}_{email_sender}_{sub_dirname}'
		datetime = dateutil.parser.parse(email_time).astimezone(dateutil.tz.tzlocal()).strftime('%d/%m/%Y %X')

		sub_dirname = os.path.join(os.getcwd(), sub_dirname)
		check_isdir(sub_dirname)
		os.chdir(sub_dirname)

		attachment_dir = os.path.join(os.getcwd(), 'atttachments')

		for part in disposition :

			filename = part.get_filename()
			filename = re.sub('[<>|?*:"/\\\\]', '', filename)

			if len(filename) > 255 : # max limit for filename
				name, ext = os.path.splitext(filename)
				name = name[:(255 - len(ext))]
				filename = name + ext


			if filename : # 'inline disposition' may or may not have an attachment

				filename = os.path.basename(filename)

				check_isdir(attachment_dir)
				os.chdir(attachment_dir)

				with open(filename, 'wb') as file :
					file.write(part.get_payload(decode=True))

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

		display_value += 1

	print('Logging out .....')
	conn.logout()
	print('\tSuccessfully logged out.\n')

	os.chdir(os.environ.get('userprofile'))

	print(f'All mails have been saved to {download_dir}')

#######################################################################################################################################

IMAP_HOST = 'imap.gmail.com'
IMAP_SSL_PORT = 993

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
