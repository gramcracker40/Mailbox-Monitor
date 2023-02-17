import win32com.client as win32
from datetime import datetime
import os
import datetime

def send_NDAemail(user):
    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.Subject = 'You signed an NDA!'
    mail.To = user

    # setting an image 
    attachment = mail.Attachments.Add(os.getcwd() + "\\logo.jpg")
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")

    time = datetime.datetime.now()

    outputF = open("mailbox_monitor_log.txt", 'a')
    outputF.write(f"NDA order confirmation responder: email sent to '{user}' at {time}\n")

    #<img src="cid:currency_img"><br><br>
    mail.HTMLBody = r"""
    <body>
    Hi,<br><br>
    
    
    Thank you for signing Web Fire’s Non-Disclosure Agreement. We look forward to working with you on all your technology needs. To help us get a better understanding of your network we ask that you follow the link below to fill out our Pre-Network scan Questionnaire.  
    This will better prepare us for scanning your network and will help in the gathering of information process. Thank you for your time.  If you have any questions, please let us know. <br>
    <a href='https://wf.myportallogin.com/tickets/categories/0e125b9c-3aa4-11ed-b8ec-020ad877b572/9b2bd233-3b60-11ed-b8ec-020ad877b572'>Portal (myportallogin.com)</a>


    <br><br>
    Sincerely,
    <br><br>
    Web Fire Communications, Inc
    <br><br>
    <img src="cid:logo">

    </body>

    """

    mail.Send()

    





