import logging
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


logging.basicConfig(filename="/home/ubuntu/reportproject/mailer/script.log", level=logging.DEBUG, format='%(asctime)s %(levelname)s: %(message)s')

def schedule_project():
    now = datetime.now()
    hours = now.strftime("%H")
    minutes = now.strftime("%M")
    val = Student.objects.filter(hour=hours, minutes=minutes).values()

    for i in val:

        val6 = i['Engineer_Name']
        print(val6)
        to_list = []
        cc_list = []
        for key, value in i.items():
            if key.startswith('sender'):
               to_list.append(value)
            elif key.startswith('cc'):
               cc_list.append(value)


        print(val6, 'wfegegegegergew')
        result = subprocess.run(['python3', "/home/ubuntu/reportproject/rameez/" + val6, *to_list,*cc_list])

        logging.info(f"Result for {val6}:")
        logging.info(f" - Output: {result.stdout.decode('utf-8').strip()}")
        logging.info(f" - Return Code: {result.returncode}")
        logging.info(f" - Result: {result}")


def hello():
    try:
         schedule_project()

    except Exception as e:
        logging.error(f"An error occurred: {e}")


if __name__ == "__main__":
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
    django.setup()
    hello()
