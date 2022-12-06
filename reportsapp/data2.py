import smtplib
import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

msg = MIMEMultipart('alternative')
msg['subject'] = 'table'
msg['from'] = 'rameeshmbabu8489@gmail.com'
msg['to'] = 'mohammed.m@futurenet.in'
text = ""

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()

