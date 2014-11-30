import os
import imaplib
from email import email
from config import USERNAME, PASSWORD, QUEUE, SUBJECTS

# Sets credentials
server = 'outlook.office365.com'
username = USERNAME
password = PASSWORD

# Establsihes connection
connection = imaplib.IMAP4_SSL(server,993)
connection.login(username,password)

# Selects Inbox
status, unread = connection.select('INBOX')

# Searches for emails based on subject
found = {}
y = 0
for subject in SUBJECTS:
    if connection.search(None, '(SUBJECT "%s")' % (subject)):
        status, num = connection.search(None, '(SUBJECT "%s")' % (subject))
        split = num[0].split()
        for item in split:
            found[item] = subject

# Assembles raw email data into human readable message
y = 0
for key in found:
    status, msg_data = connection.fetch(key, '(RFC822)')
    mail = email.message_from_string(msg_data[0][1])
    for part in mail.walk():
        if part.get_filename() != None:
            fileName = part.get_filename()
            filePath = os.path.join(QUEUE, fileName)
            if bool(fileName):
                if os.path.isfile(filePath):
                    newName = str(y) + fileName
                    newPath = os.path.join(QUEUE, newName)
                    fp = open(newPath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    y += 1
                else:
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()

# Copies downloaded attachment email to Processed folder in mailbox
myDict = dict.copy(found)
for key in found:
    connection.copy(key, 'Processed')
    status, response = connection.store(key, '+FLAGS', r'(\Deleted)')
    myDict.pop(key, None)

status, response = connection.expunge()

# Logout
connection.logout()
