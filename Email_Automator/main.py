# smpt-mail.outlook.com
import smtplib
from email.message import EmailMessage
import xlrd
import sys

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
print("successful login\n")

recipient = input("Please enter the email address of the recipient\n")

workbook = xlrd.open_workbook("templates.xlsx")
templateSheet = workbook.sheet_by_name("templates")

found = False
templateName = input("Please enter the name of the template you want to use or type quit\n")
while not found:
    if "quit" in templateName:
        sys.exit(0)
    if templateName in templateSheet.col_values(0):
        found = True
        break
    templateName = input(templateName +" was not found, please enter the name of the template you want to use or type quit\n")

email = EmailMessage()
email['From'] = username
email['To'] = recipient
#email['Subject'] = subject
#email.set_content(body)

#connection.send_message(email)



connection.quit()

