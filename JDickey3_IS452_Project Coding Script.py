
import imaplib
import email
from datetime import datetime, timedelta
from collections import Counter
import csv

ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = str(input("What is your Gmail username?")) + ORG_EMAIL
FROM_PWD    = str(input("What is your Gmail password?"))
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

mail = imaplib.IMAP4_SSL(SMTP_SERVER)
mail.login(FROM_EMAIL,FROM_PWD)
mail.select('inbox')

type, data = mail.search(None, 'ALL')
mail_ids = data[0]
id_list = mail_ids.split()
first_email_id = int(id_list[0])
latest_email_id = int(id_list[-1])

num_of_emails = int(input("How many of the latest emails should be analyzed?"))
num_of_days = int(input("How many days back should be analyzed?"))
List_of_Senders = []
allrows = []
date_N_days_ago = datetime.now(tz=None) - timedelta(days=num_of_days)

for i in range((latest_email_id-num_of_emails),latest_email_id):
    result, mail_data = mail.fetch(id_list[i],'(RFC822)')
    raw_email = mail_data[0][1].decode("utf-8")
    email_message = email.message_from_string(raw_email)
    if email_message['Date'][0:3] in ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]:
        Email_Date = email_message['Date'][0:25].strip()
        Formatted_Date= datetime.strptime(Email_Date, '%a, %d %b %Y %X')
    else:
        Email_Date = email_message['Date'][0:20].strip()
        Formatted_Date= datetime.strptime(Email_Date, '%d %b %Y %X')
    if Formatted_Date >= date_N_days_ago:
        row = []
        Date = Formatted_Date
        Sender = email_message['From']
        Subject = email_message['Subject']
        row.append(Date)
        row.append(Sender)
        row.append(Subject)
        allrows.append(row)
        List_of_Senders.append(email_message['From'])

counted = Counter(List_of_Senders)
Top10Senders = counted.most_common(10)

print("\n","\n","Analysis Results:","\n")
print("The total number of emails in your inbox is:",latest_email_id,"\n")
print("The timeframe specified is:", date_N_days_ago,"-",datetime.now(tz=None),"\n")
print("The number of emails analyzed within the specified timeframe is:",len(List_of_Senders),"\n")
print("The top ten senders followed by volumes of emails sent are as follows:","\n")

for i in range(len(Top10Senders)):
    print(i+1,". ",Top10Senders[i][0].replace("(","").replace("'","").replace(")","").replace("\"","")," : ",Top10Senders[i][1], sep='')

headers = ['Date', 'Sender', 'Subject']
with open('EmailAnalysisResults.csv', 'w', newline='') as outfile:
    csvout = csv.writer(outfile)
    csvout.writerow(headers)
    csvout.writerows(allrows)


