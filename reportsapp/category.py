import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
category_query = """SELECT  (SELECT  substring(thi.name, 31, 
position('%%OldValue%%' in thi.name) - 31) 
FROM ticket_history thi 
WHERE thi.ticket_id=t.id AND thi.name like '%FieldName%Category%'
AND thi.create_time=(SELECT max(thii.create_time) 
FROM ticket_history thii
WHERE thii.ticket_id=thi.ticket_id 
AND thii.name like '%FieldName%Category%')LIMIT 1) as Category,
SUM(CASE WHEN ts.id IN (1,4,6,7,8,11,13, 15, 16, 14, 17,18)  THEN 1 ELSE 0 END) AS 'pending',
SUM(CASE WHEN ts.id IN (2, 3,10,12) THEN 1 ELSE 0 END) AS 'Closed'
FROM ticket t
JOIN ticket_state ts ON t.ticket_state_id = ts.id
JOIN ticket_type tt ON t.type_id = tt.id
JOIN queue q ON t.queue_id = q.id
WHERE   tt.name NOT IN ('junk') AND ts.name NOT IN ('merged') 
AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK') 
AND t.create_time BETWEEN '2023-01-01' AND '2023-01-29'
group BY Category  
ORDER BY Category ASC"""
cursor.execute(category_query)
categoryresult = cursor.fetchall()

html_table = """

<html>
  <body>
    <table border="1">
      <tr>
        <th>Category</th>
        <th>pending</th>
        <th>Closed</th>
      </tr>
"""
pending_total = 0
closed_total = 0
for row in categoryresult:
    pending_total += row[1]
    closed_total += row[2]
html_table += "<tr><td><b>Totals</b></td><td><b>" + str(pending_total) + "</b></td><td><b>" + str(
    closed_total) + "</b></td></tr>"

for row in categoryresult:
    html_table += "<tr>"
    for item in row:
        html_table += "<td>" + str(item) + "</td>"
    html_table += "</tr>"
html_table += """
    </table>
  </body>
</html>
"""

message = MIMEMultipart()
message["from"] = "lokesh.p@futurenet.in"
message['to'] = "lokesh.p@futurenet.in"
message['subject'] = 'test'
body = MIMEText(html_table, 'html')
message.attach(body)

s = smtplib.SMTP(host='webmail.futurenet.in', port=587)
s.starttls()
s.login("lokesh.p@futurenet.in", "Classic@123")
s.send_message(message)
s.quit()
print("SENT")




