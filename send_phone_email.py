import win32com.client as win32
from datetime import datetime
import os
import datetime

def send_phone_email(user):
    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    mail.Subject = 'You just approved an order!'
    mail.To = user

    # setting an image 
    attachment = mail.Attachments.Add(os.getcwd() + "\\logo.jpg")
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")

    time = datetime.datetime.now()

    outputF = open("mailbox_monitor_log.txt", 'a')
    outputF.write(f"Phone LOA/questionnaire Responder: email sent to '{user}' at {time}\n")

    #<img src="cid:currency_img"><br><br>
    mail.HTMLBody = r"""
    <body>
    Hi,<br><br>
    
    
    Thank you for approving our proposal for new telephone services. Web Fire pride’s itself on delivering the highest quality customer service possible. Here is the additional paperwork that we need completed to ensure a smooth implementation of your new phone services. First, if we are porting numbers from another carrier, we need the Letter of Authorization (LOA) completed and returned to us along with your latest phone bill. Second, please complete and submit as much of the Preliminary Telephone Deployment Questionnaire as possible. If you do not have all the information for the questionnaire, we will assist with getting it. <br>
    <a href='https://wf.myportallogin.com/tickets/categories/3806504e-44c1-11ed-a2b1-020ad877b572'>Portal (myportallogin.com)</a>


    <br><br>
    Sincerely,
    <br><br>
    Web Fire Communications, Inc
    <br><br>
    <img src="cid:logo">

    </body>

    """

    mail.Send()