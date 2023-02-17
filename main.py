# Web Fire Communications, Inc
# Order Confirmation Script
# this is the script for the main process of the auto responder built for 
# checking the sales@wf.net emails sent box from a forwarded address and 
# determining if the sent email was an order confirmation of an Order Approved 
# Web Fire Non Disclosure Agreement or a Hosted Phone Service email and 
# will determine which email to send in the respective situations

# I used all standard libraries to python except pywin32 for operating outlook
import email
import imaplib
from send_NDAemail import send_NDAemail
from send_phone_email import send_phone_email

smtp_sender = ""
smtp_host = ""
smtp_pass = ""
imap_host = ""
imap_user = ""
imap_pass = ""

# To hold different parts of the email - body will be extracted in the main process
# to allow easier searching of the email
class EmailInfo:
    def __init__(self, email_full_, body_):
        self.email_full = email_full_
        self.body = body_


def find_NDA(emails):
    NDA_emails = 0
    for email in emails:
        
        if "Web Fire Non Disclosure Agreement" in f"{email.body}":
            NDA_emails = email.email_full.get('To')
            break
        
    return NDA_emails


def find_phone_confirmations(emails):
    phone_emails = 0
    for email in emails:
        
        if "Hosted Telephone Service" in f"{email.body}":
            phone_emails = email.email_full.get('To')
            break
        
    return phone_emails


def delete_mailbox():
    mail = imaplib.IMAP4_SSL(imap_host)
    #authentication
    mail.login(imap_user, imap_pass)

    mail.select('Inbox')

    _, data = mail.search(None, "ALL") #Filter by all

    for num in data[0].split():
       #deleting the emails
       mail.store(num, '+FLAGS', r'(\Deleted)')

    #parmanently deleting the emails that are selected
    mail.expunge()
    #closing the mailbox
    mail.close()
    #loging out form the mail id
    mail.logout()


####################################################################
####################################################################
#####  Runs the main process, finds all emails with subject line
#####  'Order Approved' and has "Web Fire Non Disclosure Agreement"
####  or 'Order Approved' and has "Hosted Telephone Service" 
#####  listed in the body. returns them in users_to_be_sent
####################################################################
####################################################################

# logging into IMAP, doing SSL version - check credential vars above
imap = imaplib.IMAP4_SSL(imap_host, port=993)
imap.login(imap_user, imap_pass)

#selecting the inbox to grab the emails from
imap.select("Inbox")

#grabbing all the emails in the Inbox folder
_, message_nums = imap.search(None, "ALL")

# this loop cycles through each email, adding the email to needing_action
# if its subject is an "Order Approved"
needing_action = []
for mes in message_nums[0].split():
    _, data = imap.fetch(mes, "(RFC822)")

    message = email.message_from_bytes(data[0][1])

    subj = message.get('Subject')
    
    #checking for content type, if it is multi, we have to walk to grab payload
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/plain" or part.get_content_type() == "text/html":
                body = part.get_payload(decode=True)  # decode
                break
    else:
        body = message.get_payload(decode=True)

    #Order Approved confirms that the email was a confirmation
    #pass the payload(body) and the entire message into the list of 
    #emails needing action. 
    if "Order Approved" in subj: 
        needing_action.append(EmailInfo(message, body))


#Users to be sent is determined by find_NDA and find_phone_confi - check function docs for info
nda_users_to_be_sent = find_NDA(needing_action)  # If no users need the email, users_to_be_sent = 0

phone_users_to_be_sent = find_phone_confirmations(needing_action)


#Users to be sent is any new emails found in the inbox that show a confirmation
#and signature of a NON DISCLOSURE AGREEMENT being ported back from the online signature

if nda_users_to_be_sent != 0:
    
    if len(nda_users_to_be_sent) > 1:
        users_to_be_sent = nda_users_to_be_sent.split('"')

        user = users_to_be_sent[0]
    else:
        user = nda_users_to_be_sent

    print(user)
    #send_NDAemail(user)


    

if phone_users_to_be_sent != 0:
    
    if len(phone_users_to_be_sent) > 1:
        to_be_sent = phone_users_to_be_sent.split('"')

        user = to_be_sent[0]
    else:
        user = phone_users_to_be_sent

    print(user)

    #send_phone_email(user)

    

#delete_mailbox()

imap.close()
imap.logout()





# Below code is simply leftover from the build process
# Can be used to help discover other functionality toward the script.


#print(users_to_be_sent[1])

#delete_mailbox()

# print(f"BODY:\n{body}\n\n")
    

# print(f"\n\nNEEDING ACTION:::\n\n{needing_action}")


# print(f"Message Number: 'n{mes}")
# print(f"From: \n{message.get('From')}")
# print(f"To: \n{message.get('To')}")
# print(f"Date: \n{message.get('Date')}")
# print(f"Subject: \n{message.get('Subject')}")

# print("Content:  \n\n")
# if message.is_multipart():
#     for part in message.walk():
#         if part.get_content_type() == "text/plain":
#             print(f"PART\n\n{part.as_string()}")
#             body = part.get_payload(decode=True)  # decode
#             break
#                       # not multipart - i.e. plain text, no attachments, keeping fingers crossed
# else:


# server = IMAPClient(imap_host, use_uid=True)
# server.login(imap_user, imap_pass)


# inbox = server.select_folder('INBOX')

# sales_messages = server.search(['FROM', 'sales@wf.net'])

# #Printing number of emails in inbox from sales@wf.net
# print(len(sales_messages))


# new_messages = server.search(["NOT", "DELETED"])
# response = server.fetch(new_messages, ['RFC822', 'BODY[TEXT]'])


# # body = ""
# # for uid, message_data in response.items():
# #     email_message = email.message_from_bytes(message_data[b'RFC822'])

# #     print(email_message.get_content_type())

# #     for part in email_message.walk():
# #         if part.get_content_type() == "text/plain":
# #             print(part.keys())
# #             body = email.message_from_string(message_data['BODY[TEXT]'])

# #     parsedBody = email_message.get_payload()

# #     print(uid, email_message.get("From"), email_message.get("Subject"))
# #     print(parsedBody)

# # for mail_id, data in server.fetch(new_messages,['ENVELOPE','BODY[TEXT]']).items():
# #            envelope = data[b'ENVELOPE']
# #            body = data[b'BODY[TEXT]']
# #            print(body.decode('utf-8', 'strict'))

# server.logout()


# # #Grabs ID, Subject, and date received of all emails in the inbox
# # for msg_id, data in server.fetch(sales_messages,['ENVELOPE']).items():
# #      envelope = data[b'ENVELOPE']
# #      print('ID #%d: "%s" received %s' % (msg_id, envelope.subject.decode(), envelope.date))




# send email
# def send_mail(address):
#     with smtplib.SMTP(smtp_host, 587) as smtp:
#         #encrypts traffic
#         context=ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
#         context.set_ciphers('DEFAULT@SECLEVEL=1')
        
#         smtp.starttls(context=context)
#         smtp.ehlo()

#         smtp.login(smtp_sender, smtp_pass)

#         msg = EmailMessage()
#         msg['Subject'] = "you just signed an NDA"
#         msg['From'] = smtp_sender
#         msg['To'] = address

#         msg.set_content("Here's this link 'https://web-fire-communications.itglue.com/'")

#         smtp.send_message(message)




# # # connect to host using SSL
# # imap = imaplib.IMAP4_SSL(imap_host)

# # ## login to server
# # imap.login(imap_user, imap_pass)

# imap.select('Inbox')

# tmp, data = imap.search(None, 'ALL')
# for num in data[0].split():
#         helper = contentmanager.raw_data_manager()
#         tmp, data = imap.fetch(num, '(RFC822)')
#         print('Message: {0}\n'.format(num))
#         mail = email.message_from_bytes(data[0][1])
#         header_from = decode_header(mail["From"])
#         header_to = decode_header(mail["To"])
#         date = decode_header(mail["Date"])
#         content_type = mail.get_content_type()
#         body = helper.get_content(mail, errors='replace')   #, preferencelist=('plain',)
#         #print(mail.keys())
#         #print(mail.values())
#         print(header_from)
#         print(header_to)
#         print(date)
#         print(content_type)
	
# imap.close()