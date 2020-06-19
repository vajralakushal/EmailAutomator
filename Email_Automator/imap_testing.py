
import imaplib
import pprint

imap_host = 'imap-mail.outlook.com'
imap_user = 'khandevtesting@outlook.com'
imap_pass = 'Developer2020!'

# connect to host using SSL
imap = imaplib.IMAP4_SSL(imap_host, 993)

## login to server
imap.login(imap_user, imap_pass)

imap.select('Inbox')

tmp, data = imap.search(None, 'ALL')
if tmp == "OK":
    for num in data[0].split():
        tmp, data = imap.fetch(num, '(RFC822)')
        print('Message: {0}\n'.format(num))
        pprint.pprint(data[0][1])
        break


imap.close()