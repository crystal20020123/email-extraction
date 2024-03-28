import imaplib
import email
from imap_tools import MailBox, A
# Set up the IMAP connection
mail = imaplib.IMAP4_SSL('mail.businessnbeams.com')
mail.login('info@businessnbeams.com', 'ratpiper2020')

# mail = imaplib.IMAP4_SSL('outlook.office365.com')
# mail.login('info@primedicalja.com', '2024automation')

mail.select('inbox')

# Search for all email messages in the inbox
status, data = mail.search(None, 'ALL')

# Iterate through each email message and print its contents

num = data[0].split()[-6]
status, data = mail.fetch(num, '(RFC822)')
email_message = email.message_from_bytes(data[0][1])

print('From:', email_message['From'])
print('Subject:', email_message['Subject'])
print('Date:', email_message['Date'])
print('Body:', email_message.get_payload())

flagPDF = False
if email_message.is_multipart():
     for part in email_message.walk():
        if part.get_content_type() == 'application/pdf':
            filename = part.get_filename()
            flagPDF = True
            break

else:
    if email_message.get_content_type() == 'application/pdf':
        filename = email_message.get_filename()
        flagPDF = True

if flagPDF:
    print('attachment', filename)


print()
    
# Close the connection
mail.close()
mail.logout()
