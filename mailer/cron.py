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


# def get_to_list():
#     now = datetime.now()
#     hours = now.strftime("%H")
#     minutes = now.strftime("%M")
#     # to_list = []
#     # cc_list = []
#
#     if Student.objects.filter(hour=hours, minutes=minutes).exists():
#         val = Student.objects.filter(hour=hours, minutes=minutes).values()
#         for item in val:
#             to_list = []
#             cc_list = []
#             print(to_list)
#             print(cc_list )
#
#             for key, value in item.items():
#                 if key.startswith('sender'):
#                     to_list.append(value)
#                 elif key.startswith('cc'):
#                     cc_list.append(value)
#
#
#     logging.info(f"Recipients: {to_list}, CC: {cc_list}")
#
#     return {
#         'to_list': to_list,
#         'cc_list': cc_list
#     }


def hi():
    now = datetime.now()
    hours = now.strftime("%H")
    minutes = now.strftime("%M")
    val = Student.objects.filter(hour=hours, minutes=minutes).values()

    for i in val:
        print(i)
        val6 = i['Engineer_Name']
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
        # to_list = get_to_list()
        hi()
        # logging.info(f"Recipients: {to_list['to_list']}, CC: {to_list['cc_list']}, HI: {hi}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")


if __name__ == "__main__":
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
    django.setup()
    hello()
