# Importing all the necessary modules
import os
import smtplib
import csv
from getpass import getpass
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Let's start by creating a function that reads the CSV
def read_template(filename):
    with open (filename,"r",encoding="UTF-8") as template_file:
        template_file_content = template_file.read()
        return Template(template_file_content)

# Configuring SMTP server
MY_ADDRESS = 'Enter your email addres!'
PASSWORD = getpass() # Display a box to enter the password
s = smtplib.SMTP(host='smtp.gmail.com',port = 587)
s.starttls()
s.ehlo()
s.login(MY_ADDRESS,PASSWORD)

# Read the message template
filename = '/Users/Benciowski/Documents/VisualStudio/Emailer/Setup/emailbodytemplate.txt' # Enter the absolute path to your email body template file
message_template = read_template(filename)
details_file = '/Users/Benciowski/Documents/VisualStudio/Emailer/Setup/customerdetails.csv' # Enter the absolute path to your file that contains the customer emails to email

with open (details_file,"r") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=";")
    
    # Skip the header row
    next(csv_reader)

    for row in csv_reader:
        message= ""
        msg = MIMEMultipart() # Initialiazing the message

        message = message_template.substitute(PERSON_NAME = row[0])

        # Set up the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=row[1]
        msg['Subject'] = "Thank your for ordering Genteken!" 

        # Add the body
        email_body = MIMEText(message,"html","utf-8")
        msg.attach(email_body)

        #Attach Image 
        fp = open('/Users/Benciowski/Documents/VisualStudio/Emailer/Images/Footer_Finalversion.gif', 'rb') # I want to add an image as footer
        msgImage = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<image1>') # I replace the image1 in my template per the photo I want as footer
        msg.attach(msgImage)
              
        #send the message via the server
        try:
            s.send_message(msg)
            # Confirm the email is succesfully sent!
            print(f'Email sent succesfully to: {row[1]}')

        # Catch emails in error!
        except smtplib.SMTPException:
		        print (f'Unable to send email to: {row[1]}')       
        
del msg
s.quit()

