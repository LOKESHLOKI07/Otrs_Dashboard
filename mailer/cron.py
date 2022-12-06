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


def hi():
    now = datetime.now()
    hours = now.strftime("%H")
    # print(hours)
    # hours1 = int(hours)
    # print(hours1)
    minutes = now.strftime("%M")
    val = Student.objects.all().values()
    # print(val)
    for i in val:
        # print(i['Engineer_Name'])

        hour = i['hour']
        # hour = 1
        a = pd.to_datetime(hour, format='%H')
        val1 = a.strftime("%H")
        val2 = a.strftime("%H")
        # print(val1)

        minute = i['minutes']
        a = pd.to_datetime(minute, format='%M')
        va1 = a.strftime("%M")
        va2 = a.strftime("%M")
        # print(va1)

        if val1 == hours and va1 == minutes:
            val6 = i['Engineer_Name']

            """def run_background():
                subprocess.run(["/home/ubuntu/reportproject/rameez" + val6], shell=True, stdin=None, stdout=None,
                               stderr=None, close_fds=True)
                # subprocess.run(['tertiary.py'], shell=True, stdin=None, stdout=None, stderr=None, close_fds=True)
            bg_thread = threading.Thread(target=run_background)
            bg_thread.start()"""
            # result = subprocess.run([sys.executable, "-c", "/home/ubuntu/reportproject/rameez" + val6])
            # !/usr/bin/python2.7 --> #!/usr/bin/python
            # result = subprocess.run("/home/ubuntu/reportproject/rameez/" + val6, capture_output=True, shell=True)
            # result = subprocess.run('python3',"""/home/ubuntu/reportproject/rameez/ '""" + val6 + """' """)
            # result = subprocess.run(["/home/ubuntu/reportproject/rameez", "-c", val6])
            # result = subprocess.run([sys.executable, "-c","/home/ubuntu/reportproject/rameez/" + val6])
            # result = subprocess.run(["python3", "-c","/home/ubuntu/reportproject/rameez/" + val6], capture_output=True)
            """result = subprocess.run(
                [sys.executable, "-c", "/home/ubuntu/reportproject/rameez/" + val6], capture_output=True, text=True
            )
            print(result)
            result1 = result.stdout"""
            """command = "/home/ubuntu/reportproject/rameez/" + val6
            p = subprocess.Popen(
                [command],
                shell=True,
                stdin=None,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                close_fds=True)
            print(command)
            out, err = p.communicate()"""
            # result = subprocess.run(['python3'], input="/home/ubuntu/reportproject/rameez/" + val6, capture_output=True, encoding='UTF-8')
            result = subprocess.run(['python3', "/home/ubuntu/reportproject/rameez/" + val6])
            # result2 = result.stderr
            #print(result)
            # print(result2)

            #print(result.stdout)
        else:
            print('error')


hi()
