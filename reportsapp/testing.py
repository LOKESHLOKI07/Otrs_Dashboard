from datetime import datetime, timedelta

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
  SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18) and t.create_time  BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS pending,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.create_time BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS closed,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.change_time BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS 'overallClosed'
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
engineer_query = """c
 """
cursor.execute(engineer_query)
engineerresult = cursor.fetchall()

category_query = """SELECT 
 (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
  SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18) and t.create_time BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS pending,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.create_time BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS closed,
SUM(CASE WHEN ts.id IN (2, 3,10,12) and t.change_time BETWEEN DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 MONTH) 
and DATE_SUB(DATE_SUB(LAST_DAY(CURRENT_DATE), INTERVAL DAY(LAST_DAY(CURRENT_DATE)) - 1 DAY), INTERVAL 1 DAY)   THEN 1 ELSE 0 END) AS 'overallClosed'
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
today = datetime.now().strftime("%d/%m/%Y")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
html_table += f"<p style='font-size: 20px;'> Dear team, <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find the below statistics of Daily Registered ticket Summary from {yesterday} 06:00 AM to {today} 06:00 AM </p>"

result_list = [typeresult, categoryresult, engineerresult]

for i in typeresult:
    pending_total = 0
    closed_total = 0
    total = 0

    for row in typeresult:
        pending_total += row[1]
        closed_total += row[2]
total = pending_total + closed_total
print(total)
# html_table += f"<tr><td style='text-align: left; width: 100px; '><b>Overall Totals</b></td><td style='text-align: center; width: 100px; '><b>{total}</b></td></tr>"

for result, result_name in zip(result_list, ["Types", "Category", "Engineer"]):

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
        # total = 0
        # final_tot = []
        for row in result:
            pending_total += row[1]
            closed_total += row[2]
        #     total = pending_total + closed_total
        # final_tot.append(total)

        html_table += f"<tr><td style='text-align: left; width: 100px; '><b>Total</b></td><td style='text-align: center;  width: 100px;'><b>{total}</b></td><td style='text-align: center;  width: 100px;'><b>{closed_total}</b><td style='text-align: center;  width: 100px;'><b>-</b></td></td><td style='text-align: center; width: 100px; '><b>{pending_total}</b></td></tr>"

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
today = datetime.now().strftime("%d/%m/%Y")
yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

message = MIMEMultipart()
message["from"] = "otrs.report@futurenet.in"
message['to'] = "lokesh.p@futurenet.in"
message['subject'] = "Daily Registered Ticket Summary" + yesterday + ' 6:00 AM to ' + today + ' 6:00 AM '
body = MIMEText(html_table, 'html')
message.attach(body)

s = smtplib.SMTP(host='webmail.futurenet.in', port=587)
s.starttls()
s.login("otrs.report@futurenet.in", "WElcome@123")
s.send_message(message)
s.quit()
print("SENT")
# gopinath@futurenet.in
