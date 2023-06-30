from django.shortcuts import render

#import required libs
import pandas as pd
import smtplib,time
import os
#import required model
from mailapp.models import DATA
# Create your views here.
def index(request):
    if (request.method=='POST'):
        senders_mail=request.POST['senders_mail']
        app_pass=request.POST['app_pass']
        senders_name=request.POST['senders_name']
        message=request.POST['message']
        subject=request.POST['subject']
        file=request.FILES['file']
        ins=DATA(location=file)
        ins.save()
        file_path = 'media/'+ ins.location.name

        logs=send_email_static(senders_mail,app_pass,senders_name,message,subject,file_path)
        #delete the file from local storage 
        os.remove(file_path)

        context={
            'log':logs
        }

        return render(request,'success.html',context)

        


    return render(request,'home.html')


def send_email_static(senders_mail,app_pass,senders_name,message,subject,file_path):
    logs=[]
    #establishing connection
    server = smtplib.SMTP_SSL('smtp.gmail.com',465)
    server.ehlo()
    server.login(senders_mail,app_pass)

    # Read the file entered by user
    email_list = pd.read_excel(file_path) 

    # Get all the Names, Email Addreses, Subjects and Messages
    all_names = email_list['NAME']
    all_emails = email_list['EMAIL']
    
    # Loop through the emails
    for idx in range(len(all_emails)):
        # Get each records name, email, subject and message
        name = all_names[idx]
        email = all_emails[idx]

        msg=f"Dear {name},\n"


        # Create the email to send
        full_email = ("From: {0} <{1}>\n"
                      "To: {2} <{3}>\n"
                      "Subject: {4}\n\n"
                      "{5}{6}"
                      .format(senders_name, senders_mail, name, email, subject, msg, message))#name change

        # In the email field, you can add multiple other emails if you want
        # all of them to receive the same text
        try:
            server.sendmail(senders_mail, [email], full_email)
            logs.append(f'Email to {name}->{email} successfully sent!')
            #print(f'Email to {name}->{email} successfully sent!')
            time.sleep(1)#important->adds delay of 1 sec after every mail making rate 60 mails/min since smtp limit is 80mails/min
        except Exception as e:
            logs.append('Email to {} could not be sent--------------------------------------------------------------because {}'.format(name, str(e)))

    # Close the smtp server
    server.close()
    return logs

