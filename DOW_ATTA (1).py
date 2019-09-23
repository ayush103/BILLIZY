import imaplib,email
user=raw_input('Enter gmail id: ')
password=raw_input('Enter password: ')
imap_url = 'imap.gmail.com'

attachment_dir = 'downloaded_attachments'
import os
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

mail = imaplib.IMAP4_SSL(imap_url)
mail.login(user, password)
mail.list()
# Out: list of "folders" aka labels in gmail.
mail.select("inbox") # connect to inbox.

# -------------------------------------------- #
# allows you to download attachments
def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))


result, data = mail.uid('search', None, "ALL") # search and return uids instead
last_50_ids = data[0].split()[-10:-1]#-50:0][::-1]
latest_id = data[0].split()[-1]
last_50_ids.append(latest_id)
# print(last_50_ids)
last_50_ids=last_50_ids[::-1]
c=1
for i in last_50_ids:
    print ('Fetching mail:',c)
    result, data = mail.uid('fetch', i, '(RFC822)')
    raw_email = data[0][1]
    # print(raw_email)
    # --------------------- formatting ------------ #
    email_message = email.message_from_string(raw_email)
    print ('\tTo',email_message['To'])
    print ('\tFrom',email.utils.parseaddr(email_message['From']) )# for parsing "Yuji Tomita" <yuji@grovemade.com>
    # print (email_message.items()) # print all headers
     # --------------------------------------------------- #
    get_attachments(email_message)
    print('\tDowloaded attachments for this mail !')
    c+=1
