

import imaplib
import email

ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = "dickey.john" + ORG_EMAIL
FROM_PWD    = "XXXXX"
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

print(first_email_id)
print(latest_email_id)
print(id_list[0])

result, mail_data = mail.fetch(id_list[0],'(RFC822)')
# print(result)
# print(mail_data)

def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)

def search(key, value, mail):
    result, data = mail.search(None, key,'"{}"'.format(value))
    return data


raw = email.message_from_bytes(mail_data[0][1])
# print(get_body(raw))
print(search('FROM',"offers@wish.com",mail))
# for i in range(latest_email_id,first_email_id, -1):
#     typ, data = mail.fetch(i, '(RFC822)' )

