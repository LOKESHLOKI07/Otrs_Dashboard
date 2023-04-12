import sys
import os
import django

"""sys.path.append('/home/ubuntu/reportproject')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
django.setup()"""
#from subprocess import run

import smtplib
import pandas as pd
from datetime import date
from datetime import datetime
from datetime import timedelta
import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys


msg = MIMEMultipart('alternative')
msg['subject'] = 'hello i am page 3 '
msg['from'] = 'lokesh.p@futurenet.in'
to_list = sys.argv[1:5]
cc_list = sys.argv[5:]
msg['bcc'] = 'lokesh.p@futurenet.in'

msg['to'] = ', '.join(to_list)
msg['cc'] = ', '.join(cc_list)

rcpt = msg['cc'].split(",") + msg['to'].split(",")
print(rcpt)

mail = smtplib.SMTP('webmail.futurenet.in', 587)
mail.ehlo()
mail.starttls()
mail.login("kevinsanjaynelthropp@futurenet.in", "KevinChris@16")
mail.sendmail("lokesh.p@futurenet.in", rcpt, msg.as_string())
mail.quit()
