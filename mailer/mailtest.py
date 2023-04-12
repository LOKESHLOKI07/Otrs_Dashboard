import os
import datetime
import smtplib


def send_email():
    sender_email = "lokesh.p@futurenet.in"
    receiver_email = "lokesh.p@futurenet.in"
    password = "Lokesh@123"
    message = "Subject: Hi\n\nThis is Lokesh"

    with smtplib.SMTP(host='webmail.futurenet.in', port=587) as smtp:
        smtp.starttls() # Add TLS encryption
        smtp.login(sender_email, password)
        smtp.sendmail(sender_email, receiver_email, message)

    # Create a text file to indicate that the email has been sent
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")
    with open("/home/ubuntu/reportproject/mailer/email_sent.txt", "a") as f:
        f.write("Email sent successfully at " + timestamp + ".\n")

send_email() # Call the function to send the email and create the text file
