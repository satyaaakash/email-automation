import pandas as pd
import smtplib
from email import policy
from email.parser import BytesParser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load the email list
email_list = pd.read_csv('opsemail.csv')

# Email credentials
sender_email = "s.obellaneni@ufl.edu"
sender_password = "Satya@2028"
smtp_server = "smtp-mail.outlook.com"
smtp_port = 587

# Read the email template
with open('ops.emltpl', 'rb') as f:
    template = BytesParser(policy=policy.default).parse(f)

# Set up the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, sender_password)

for index, row in email_list.iterrows():
    receiver_email = row['Email']  # Assuming the email column is named 'email'

    # Create a copy of the template
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = template['Subject']

    # Add the HTML body
    message.attach(MIMEText(template.get_body().get_content(), 'html'))

    # Add attachments if any
    for part in template.iter_attachments():
        message.attach(part)

    # Send the email
    server.sendmail(sender_email, receiver_email, message.as_string())

server.quit()
print("Emails have been sent successfully.")
