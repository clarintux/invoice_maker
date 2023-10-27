#!/usr/bin/env python

########################################################
# send_mail module: Please change it accordently to    #
#                   your email server and your liking. #
#                   Do not expect that everything works#
#                   without your changes!              #
#                                                      #
# Author: Clarintux                                    #
# Copyright: (C) 2023                                  #
# License: GPL-3                                       #
########################################################

import email, smtplib, ssl

from email import encoders
from email.utils import formatdate
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import pickle
import sys
import getpass


def send(customerList, password):
    """Send customed email to each customer."""
    day = datetime.datetime.today().strftime('%d')
    month = datetime.datetime.today().strftime('%B')
    year = datetime.datetime.today().strftime('%Y')

    #######################################################################
    # The user has to modify stmp_server, smtp_port,                      #
    # subject, body and filename for the attachment                       #
    #                                                                     #
    smtp_server = "smtp@server.com"                                       #
    smtp_port = 25                                                        #
    subject = "My Invoice {}".format(day + " " + month + " " + year)      #
    filename = "Invoice " + day + "_" + month + "_" + year                #
    #######################################################################

    # Load your email address
    try:
        with open("./my_business/my_infos.pkl", "rb") as myInfoFile:
            myInfos = pickle.load(myInfoFile)
        myMail = myInfos['mail']
    except:
        sys.exit("ERROR: Couldn't read file './my_business/my_infos.pkl'  :-(")

    if not password:
        password = getpass.getpass(prompt = 'Plese enter your email password: ')

    try:
        # Log in to server using secure context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()  # Can be omitted
            server.starttls(context=context)
            server.ehlo()  # Can be omitted
            server.login(myMail, password)

            for customer in customerList:
                customerName = customer['name']
                customerMail = customer['mail']
                invoicePDFName = customer['pdf']
                body = """\
Dear {}

I send you my invoice. Thanks!
This mail was automatically generated and sended.

Best regards,
INSERT YOUR NAME
                """.format(customerName)

                # Create a multipart message and set headers
                message = MIMEMultipart()
                message["From"] = myMail
                message["To"] = customerMail
                message['Date'] = formatdate(localtime=True)
                message["Subject"] = subject
                message["Bcc"] = myMail  # To receive a copy of sended emails
            
                # Add body to email
                message.attach(MIMEText(body, "plain"))
                
                # Open PDF file in binary mode
                with open(invoicePDFName, "rb") as attachment:
                    # Add file as application/octet-stream
                    # Email client can usually download this automatically as attachment
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                
                # Encode file in ASCII characters to send by email
                encoders.encode_base64(part)

                # Add header as key/value pair to attachment part
                part.add_header(
                    'Content-Disposition',
                    'attachment; filename={}'.format(filename)
                )

                # Add attachment to message and convert message to string
                message.attach(part)
                text = message.as_string()
                server.sendmail(myMail, customerMail, text)

    except:
        sys.exit("ERROR: Couldn't connect to smtp server or send emails  :-(")
