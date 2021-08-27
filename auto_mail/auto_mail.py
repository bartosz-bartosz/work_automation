# -*- coding: utf-8 -*-
"""
Created on Wed Aug 18 13:00:09 2021

I have 100 emails to send... They are almost identical,
the ony differences are title and attached file. This code does it for me.

This version uses fake addresses and password.

"""

import yagmail, os, time

user = 'email_sender@mail.com'
app_password = 'app_password_string'
to = ['email_receiver1@mail.com', 'email_receiver2@mail.com']

path = 'C:/Users/dalecb/Desktop/Invoices/'
folders = os.listdir(path)

for customer in folders:
    customer_invoices = os.listdir(path + customer)
    customer_invoices = [path + customer + '/' + file for file in customer_invoices]
    subject = customer + ' Invoices'
    content = ''
    with yagmail.SMTP(user, app_password) as yag:
        yag.send(to, subject, content, attachments=customer_invoices)
    print('Email successfully sent with title ' + subject)
    time.sleep(20)