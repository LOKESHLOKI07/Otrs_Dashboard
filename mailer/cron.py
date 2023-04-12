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

from datetime import datetime
import subprocess
import logging


def schedule_project():
    # Get the current time and day
    now = datetime.now()
    hours = now.strftime("%H")
    minutes = now.strftime("%M")
    days = now.strftime('%A')

    # Retrieve all the data that matches the given conditions
    all_val = list(Student.objects.filter(hour=hours, minutes=minutes, days__contains=days).values())
    print(all_val)
    # Loop through each data instance
    for data in all_val:
        # Retrieve the engineer name
        engineer_name = data['Engineer_Name']


        # Retrieve the email recipients and cc list
        to_list = [value for key, value in data.items() if key.startswith('sender')]
        cc_list = [value for key, value in data.items() if key.startswith('cc')]


        # Execute the Python script for the engineer with the email recipients and cc list as arguments
        subprocess.run(['python3', f"/home/ubuntu/reportproject/rameez/{engineer_name}.py", *to_list, *cc_list])



def hello():
    try:
         schedule_project()

    except Exception as e:
        logging.error(f"An error occurred: {e}")


if __name__ == "__main__":
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
    django.setup()
    hello()


