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
print(prompt_dict)

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

