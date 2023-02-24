import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
sql_query1 = "select login from users where valid_id =1 order by login asc"
cursor.execute(sql_query1)
result = cursor.fetchall()
# print(result)
