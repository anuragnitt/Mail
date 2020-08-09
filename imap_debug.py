import imaplib
import email

host = 'imap.gmail.com'
port = 993

username = input('Gmail ID : ')
password = input('Password : ')

conn = imaplib.IMAP4_SSL(host, port)
conn.login(username, password)
conn.select('INBOX')

b = conn.uid('search', None, 'ALL')[1][0].split()[::-1]

n = int(input('Faulty mail no. : '))
b = b[n-1]
print('\n'+'*'*70+'\n')

data = email.message_from_bytes(conn.uid('fetch', b, '(RFC822)')[1][0][1])

p = [x for x in data.walk()]

info = [('SUBJECT', [x['Subject'] for x in data.walk()]), ('FROM', [x['From'] for x in data.walk()]),
	   ('TO', [x['To'] for x in data.walk()]), ('DATE', [x['Date'] for x in data.walk()]),
	   ('DISPOSITION', [x['Content-Disposition'] for x in data.walk()]), ('FILENAME', [x.get_filename() for x in data.walk()])] 

print('\n DATA PARTS :\n')
for x in p :
	print(x.get_content_type())
print('='*70+'\n')

print('\nINFO TABLE :\n')
for x in info :
	print(x[0], ':', x[1], '\n')
print('='*70+'\n')