# smpt-mail.outlook.com

import smtplib
email_clients = {
    "gmail" : ("smtp.gmail.com", 587),
    "outlook" : ("smtp-mail.outlook.com", 587),
    "hotmail" : ("smtp-mail.outlook.com", 587),
    "yahoo" : ("smtp.mail.yahoo.com", 587)
}
client = input("What email client do you use? Choose from the options below\ngmail, outlook/hotmail, or yahoo?")
if client not in email_clients:
    client = input("Sorry, we don't support that email client yet. Please use lowercase letters, and make sure there aren't any typos.\n\nWhat email client do you use? Choose from the options below\ngmail, outlook/hotmail, or yahoo?")
connection = smtplib.SMTP(email_clients[client][0], email_clients[client][1])
connection.ehlo()
connection.starttls()
username = input("What's your username (i.e. your email address)?")
password = input("What's your password? If you are using gmail, please don't forget to use your app-specific password instead of your actual password")
connection.login(username, password)

recipients = []
recipient = ""
while True:
    recipient = input("Who all would you like to send your email to? Type their address in here. Type 'quit' once you're finished.")
    if "quit" in recipient:
        break
    else:
        recipients.append(recipient)


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

for recipient in recipients:
    connection.sendmail(username, recipient, email_arg)

connection.quit()

