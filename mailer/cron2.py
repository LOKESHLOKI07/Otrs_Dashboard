import smtplib
import subprocess
import threading

import pandas as pd
from datetime import date
from datetime import datetime
from datetime import timedelta
import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import calendar

import sys
import os
import django
from subprocess import run

sys.path.append('/home/ubuntu/reportproject')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
django.setup()

from mailer.models import Student
from datetime import datetime
from django.core.mail import EmailMessage

def get_to_list():
    now = datetime.now()
    hours = now.strftime("%H")
    minutes = now.strftime("%M")
    to_list = []
    cc_list = []


    if Student.objects.filter(hour=hours, minutes=minutes).exists():
        val = Student.objects.filter(hour=hours, minutes=minutes).values()
        for item in val:

            for key, value in item.items():
                if key.startswith('sender'):
                    to_list.append(value)
                elif key.startswith('cc'):
                    cc_list.append(value)
    print(to_list)
    print(cc_list)




    return {
        'to_list': to_list,
        'cc_list': cc_list
    }

def hi():

    now = datetime.now()
    hours = now.strftime("%H")
    minutes = now.strftime("%M")
    val = Student.objects.filter(hour=hours, minutes=minutes).values()


    for i in val:
        val6 = i['Engineer_Name']
        subprocess.run(['python3', "/home/ubuntu/reportproject/rameez/" + val6])

if __name__ == "__main__":
    to_list = get_to_list()
    if to_list:

        pass
    hi = hi()