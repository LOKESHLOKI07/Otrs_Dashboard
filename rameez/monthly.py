import sys
from datetime import datetime, timedelta, date

import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import time

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
type_query = """
 SELECT  tt.name,
  SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS pending,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS closed,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.change_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS 'overallClosed'
FROM 
  ticket AS t 
left JOIN ticket_state ts ON t.ticket_state_id = ts.id
left JOIN users u ON t.user_id = u.id
left JOIN ticket_priority tp ON t.ticket_priority_id = tp.id
left JOIN ticket_type tt ON t.type_id = tt.id
left JOIN queue q ON t.queue_id = q.id
WHERE 

 ts.name NOT IN ('merged') 
  AND q.name NOT IN ('SALES', 'PRESALES', 'ODOOHELPDESK') 
GROUP BY tt.name 
HAVING (pending > 0 OR closed > 0 OR overallClosed > 0)
"""
cursor.execute(type_query)
typeresult = cursor.fetchall()
engineer_query = """SELECT u.login AS Engineer_Name,
  SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS pending,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS closed,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.change_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS 'overallClosed'
FROM ticket t
left JOIN ticket_state ts ON t.ticket_state_id = ts.id
left JOIN users u ON t.user_id = u.id
left JOIN ticket_type tt ON t.type_id = tt.id
left JOIN queue q ON t.queue_id = q.id
WHERE ts.name NOT IN ('merged') 
  AND q.name NOT IN ('SALES', 'PRESALES', 'ODOOHELPDESK') 
GROUP BY u.login 
HAVING (pending > 0 OR closed > 0 OR overallClosed > 0)
ORDER BY u.login asc
 """
cursor.execute(engineer_query)
engineerresult = cursor.fetchall()

category_query = """SELECT 
 (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
  SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS pending,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS closed,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.change_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 month) AND NOW() THEN 1 ELSE 0 END) AS 'overallClosed'
FROM 
  ticket AS t 
left JOIN ticket_state ts ON t.ticket_state_id = ts.id
left JOIN users u ON t.user_id = u.id
left JOIN ticket_priority tp ON t.ticket_priority_id = tp.id
left JOIN ticket_type tt ON t.type_id = tt.id
left JOIN queue q ON t.queue_id = q.id
WHERE ts.name NOT IN ('merged') 
  AND q.name NOT IN ('SALES', 'PRESALES', 'ODOOHELPDESK') 
GROUP BY Category
HAVING (pending > 0 OR closed > 0 OR overallClosed > 0)
"""
cursor.execute(category_query)
categoryresult = cursor.fetchall()

html_table = ""

def get_first_and_last_day_of_prev_month():
    today = date.today()
    first_day_of_current_month = date(today.year, today.month, 1)
    last_day_of_prev_month = first_day_of_current_month - timedelta(days=1)
    first_day_of_prev_month = date(last_day_of_prev_month.year, last_day_of_prev_month.month, 1)
    return (first_day_of_prev_month, last_day_of_prev_month)


first_day_of_prev_month, last_day_of_prev_month = get_first_and_last_day_of_prev_month()
html_table += f"<p style='font-size: 20px;'> Dear team, <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find the below statistics of Monthly Registered ticket summary from {first_day_of_prev_month} 06:00 AM to {last_day_of_prev_month} 06:00 PM </p>"
result_list = [typeresult, categoryresult, engineerresult]

for result, result_name in zip(result_list, ["Types", "Category", "Engineer"]):
    pending_total = 0
    closed_total = 0
    overall_total = 0
    closed_total1 = 0
    total = 0
    closetotal = []
    maxtotal = []
    for row in result:
        pending_total += row[1]
        closed_total += row[2]
        overall_total += row[3] - row[2]
        closed_total1 = row[2]
        sum_value = row[1] + row[2]
        maxtotal.append(sum_value)
        closetotal.append(closed_total1)
        if row[0] == 'JUNK':
            print(row)
            junk = (row[2])
            junk1 = (row[3])
            print(junk1)

    maxclosed = max(closetotal)
    maximum = max(maxtotal)
    total = pending_total + closed_total
    maximumtotal = (total - junk)
    maxclosedtotal = (closed_total - junk1)
    # print(maximumtotal)

    if result:
        html_table += f"<h3 style='text-align: center;'>{result_name}</h3>"
        html_table += """
        <table border="1" style="border-collapse: collapse; margin: 0 auto;">

          <tr style="background-color: lightgray;">
            <th style="padding: 8px; text-align: left; border: 1px solid black;">{}</th>
            <th style="padding: 8px; text-align: center; border: 1px solid black;">Registered</th>

            <th style="padding: 8px; text-align: center; border: 1px solid black;">Closed</th>
            <th style="padding: 8px; text-align: center; border: 1px solid black;">Overall Closed</th>
              <th style="padding: 8px; text-align: center; border: 1px solid black;">Pending</th>
          </tr>
        """.format(result_name)
        pending_total = 0
        closed_total = 0

        for row in result:
            pending_total += row[1]
            closed_total += row[2]

        html_table += f"<tr><td style='text-align: left; width: 100px; '><b>Total</b></td><td style='text-align: center;  width: 100px;'><b>{maximumtotal}</b></td><td style='text-align: center;  width: 100px;'><b>{maxclosedtotal}</b><td style='text-align: center;  width: 100px;'><b>{overall_total}</b></td></td><td style='text-align: center; width: 100px; '><b>{pending_total}</b></td></tr>"

        for row in result:
            html_table += "<tr>"
            html_table += f"<td style='padding: 8px; text-align: left; border: 1px solid black; width: 50px;'>{row[0]}</td>"
            html_table += f"<td style='padding: 8px; text-align: left; border: 1px solid black; width: 50px;'>{row[1] + row[2]}</td>"

            html_table += f"<td style='padding: 4px; text-align: center; border: 1px solid black; width: 50px;'>{row[2]}</td>"
            html_table += f"<td style='padding: 4px; text-align: center; border: 1px solid black; width: 50px;'>{row[3] - row[2]}</td>"
            html_table += f"<td style='padding: 4px; text-align: center; border: 1px solid black; width: 50px;'>{row[1]}</td>"
            html_table += "</tr>"

        html_table += "</table>"

html_table += f"<p style='text-align: center; border: 1px solid black;'><b>Note</b> : This is an Auto Generated Email</p>"
html_table = f"""
<html>
  <head>
    <style>
      .table-container {{
        max-width: 800px;
        margin: 0 auto;
      }}

      table {{
        border-collapse: collapse;
        width: 100%;
      }}

      th, td {{
        padding: 8px;
        text-align: left;
        border: 1px solid black;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }}

      th {{
        background-color: lightgray;
      }}

      h3 {{
        text-align: center;
      }}
    </style>
  </head>
  <body>
    <div class="table-container">
      {html_table}
    </div>
  </body>
</html>

"""

first_day_of_prev_month, last_day_of_prev_month = get_first_and_last_day_of_prev_month()
first_day_of_prev_month_str = first_day_of_prev_month.strftime("%Y-%m-%d")
last_day_of_prev_month_str = last_day_of_prev_month.strftime("%Y-%m-%d")


# Create message object
message = MIMEMultipart()

# Set recipients
to_list = sys.argv[1:5]
cc_list = sys.argv[5:]
message['to'] = ', '.join(to_list)
message['cc'] = ', '.join(cc_list)

# Combine recipients for sending
rcpt = message['cc'].split(",") + message['to'].split(",")

# Connect to mail server and send email
with smtplib.SMTP('webmail.futurenet.in', 587) as mail:
    mail.ehlo()
    mail.starttls()
    mail.login("otrs.report@futurenet.in", "JwD@!3j@4HQB!@")
    message['subject'] = "Monthly Registered Ticket Summary " + first_day_of_prev_month_str + ' 06:00 AM to ' + last_day_of_prev_month_str + ' 06:00 AM'
    body = MIMEText(html_table, 'html')
    message.attach(body)
    mail.sendmail("otrs.report@futurenet.in", rcpt, message.as_string())

print("SENT")


