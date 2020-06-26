# smpt-mail.outlook.com
#script process: 
#read inbox -> parse each email -> put email information into excel sheet -> read excel sheet 
# -> extrapolate information -> send email to customer 
import imaplib
import base64
import os
import email
import xlrd
import openpyxl
from email.message import EmailMessage

imap_clients = {
    "outlook" : ("imap-mail.outlook.com", 993)
}

loginFile = open("credentials.txt", "r")
connectionInformation = {}
for line in loginFile:
    key = line.split(":")[0]
    value = line.split(":")[1].rstrip("\n")
    connectionInformation[key] = value

client = connectionInformation["client"]

#Connect and login to outlook mailbox
mailbox = imaplib.IMAP4_SSL(imap_clients[client][0], imap_clients[client][1])
print("successful connection to outlook\n")
username = connectionInformation["username"]
password = connectionInformation["password"]
mailbox.login(username, password)
print("successful login\n")

#Set the current mailbox to search from and start going thru unread emails
mailbox.select('Inbox')
#Container to hold unread emails
unread_emails = []
type, data = mailbox.search(None, '(UNSEEN)')
mail_ids = data[0]
id_list = mail_ids.split()
temp = data[0].split()
if type == "OK":
    for num in temp:
        typ, data = mailbox.fetch(num, '(RFC822)' )
        raw_email = data[0][1]
        for response_part in data:
            if isinstance(response_part, tuple):
                #Create new Email Message object and add it to unread_emails array
                msg = email.message_from_string(response_part[1].decode('utf-8'))

                unread_email = EmailMessage()
                unread_email['Subject'] = msg['subject']
                unread_email['From'] = msg['from']
                unread_email.set_content(msg.get_payload()[0].get_payload())

                unread_emails.append(unread_email)
print(f"Unread emails fetched:{len(unread_emails)}")

#Close connection to outlook and logout
mailbox.close()
mailbox.logout()
print("Closing connection and logging off")

workbook = openpyxl.load_workbook("templates.xlsx")
customers = workbook["Customers"]

