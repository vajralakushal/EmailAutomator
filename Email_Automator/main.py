# smpt-mail.outlook.com
import smtplib
from email.message import EmailMessage

email_clients = {
    "gmail" : ("smtp.gmail.com", 587),
    "outlook" : ("smtp-mail.outlook.com", 587),
    "hotmail" : ("smtp-mail.outlook.com", 587),
    "yahoo" : ("smtp.mail.yahoo.com", 587)
}

loginFile = open("credentials.txt", "r")
connectionInformation = {}
for line in loginFile:
    key = line.split(":")[0]
    value = line.split(":")[1].rstrip("\n")
    connectionInformation[key] = value

client = connectionInformation["client"]
connection = smtplib.SMTP(email_clients[client][0], email_clients[client][1])
connection.ehlo()
connection.starttls()
print("successful connection to outlook\n")

username = connectionInformation["username"]
password = connectionInformation["password"]
connection.login(username, password)
print("successful login")

recipients = []
recipient = ""
while True:
    recipient = input("Who all would you like to send your email to? Type their address in here. Type 'quit' once you're finished.")
    if "quit" in recipient:
        break
    else:
        recipients.append(recipient)

recipient = input("Who all would you like to send your email to? Type their address in here. Type 'quit' once you're finished.")

subject = input("What's the subject line of your email? Type redo if you want to type it again.")
while True:
    if "redo" in subject:
        subject = input("What's the subject line of your email? Type redo if you want to type it again.")
    if "redo" not in subject:
        break

body = input("What's the body of your email? Type redo if you want to type it again.")
while True:
    if "redo" in body:
        subject = input("What's the body of your email? Type redo if you want to type it again.")
    if "redo" not in body:
        break

email_arg = subject, "\n\n", body

#email = EmailMessage()
#email['From'] = username
#email['To'] = recipient
#email['Subject'] = subject
#email.set_content(body)

#print(username, recipient, subject, body)

#connection.send_message(email)

for recipient in recipients:
    print(recipient)
    connection.sendmail(username, recipient, email_arg)

connection.quit()

