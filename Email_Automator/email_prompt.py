# smpt-mail.outlook.com
#script process: 
#read inbox -> parse each email -> put email information into excel sheet -> read excel sheet 
# -> extrapolate information -> send email to customer 
import smtplib
import email
from email.message import EmailMessage
import xlrd
import sys
import imaplib

email_clients = {
    "gmail" : ("smtp.gmail.com", 587),
    "outlook" : ("smtp-mail.outlook.com", 587),
    "hotmail" : ("smtp-mail.outlook.com", 587),
    "yahoo" : ("smtp.mail.yahoo.com", 587)
}

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

mailbox = imaplib.IMAP4_SSL(imap_clients[client][0], imap_clients[client][1])

connection = smtplib.SMTP(email_clients[client][0], email_clients[client][1])
connection.ehlo()
connection.starttls()
print("successful connection to outlook\n")

username = connectionInformation["username"]
password = connectionInformation["password"]

connection.login(username, password)
mailbox.login(username, password)
print("successful login\n")


mailbox.select("INBOX")
unreadEmails=0
(retcode, messages) = mailbox.search(None, '(UNSEEN)')
#print("This is being printed by line 47 " + retcode, messages)
#print(messages[0].split())
if retcode == 'OK':
   for num in messages[0].split() :
      print ('Processing ')
      unreadEmails=unreadEmails+1
      typ, data = mailbox.fetch(num,'(RFC822)')
      #print(data)
      for response_part in data:
        if isinstance(response_part, tuple):
            original = email.message_from_bytes(data[0][1])
            
            print (original['From'])
            print (original['Subject'])
            if original.is_multipart():
                for payload in original.get_payload():
                    print(payload.get_payload())
            #else:
                #print(original.get_payload())
            
             
            typ, data = mailbox.store(num,'+FLAGS','\\Seen')

print (unreadEmails)


connection.quit()
mailbox.close()
mailbox.logout()
sys.exit(0)

recipient = input("Please enter the email address of the recipient\n")

workbook = xlrd.open_workbook("templates.xlsx")
templateSheet = workbook.sheet_by_name("templates")

found = False
template_name_row = 0
templateName = input("Please enter the name of the template you want to use or type quit\n")
while not found:
    if "quit" in templateName:
        sys.exit(0)
    if templateName in templateSheet.col_values(0):
        template_name_row = templateSheet.col_values(0).index(templateName)
        found = True
        break
    templateName = input(templateName +" was not found, please enter the name of the template you want to use or type quit\n")

prompt_col = templateSheet.row_values(0).index("Required Prompts")

prompts = templateSheet.cell_value(template_name_row, prompt_col).split(",")
prompt_dict = {}
for prompt in prompts:
    test = input("Please enter value for " + prompt + "\n")
    prompt_dict[prompt] = test

message_index = templateSheet.row_values(0).index("Template Message")
body = templateSheet.cell_value(template_name_row, message_index)
for item in prompt_dict:
    body = body.replace(item, prompt_dict[item])

subject_index = templateSheet.row_values(0).index("Subject")
subject = templateSheet.cell_value(template_name_row, subject_index)
for item in prompt_dict:
    subject = subject.replace(item, prompt_dict[item])



email = EmailMessage()
email['From'] = username
email['To'] = recipient
email['Subject'] = subject
email.set_content(body)

connection.send_message(email)
print("Email was successfully sent to ", recipient)
connection.quit()
mailbox.close()
mailbox.logout()
