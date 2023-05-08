import sys

import mysql.connector
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib


user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host=host, port=port, user=user, password=password, database=database)
cursor = db_conn.cursor()

query = """
SELECT (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,
  (SELECT SUBSTRING(CAST(thi.name AS CHAR(100)), 31, POSITION('%%OldValue%%' IN thi.name) - 31) AS customer
   FROM ticket_history thi
   WHERE thi.ticket_id = t.id
     AND thi.name LIKE '%FieldName%Category%'
     AND thi.create_time = (SELECT MAX(thii.create_time)
                            FROM ticket_history thii
                            WHERE thii.ticket_id = thi.ticket_id AND thii.name LIKE '%FieldName%Category%')
   LIMIT 1) AS category,
SUM(CASE
        WHEN ts.id IN (1, 4, 6, 7, 8, 11, 13, 15, 16, 14, 17, 18, 2, 3, 10, 12) THEN 1
        ELSE 0
      END) AS registered
FROM ticket AS t
  LEFT JOIN ticket_state ts ON t.ticket_state_id = ts.id
  LEFT JOIN users u ON t.user_id = u.id
  LEFT JOIN ticket_priority tp ON t.ticket_priority_id = tp.id
  LEFT JOIN ticket_type tt ON t.type_id = tt.id
  LEFT JOIN queue q ON t.queue_id = q.id
WHERE ts.name NOT IN ('merged')
  AND q.name NOT IN ('SALES', 'PRESALES', 'ODOOHELPDESK')
  AND t.create_time BETWEEN DATE_SUB(NOW(), INTERVAL 1 day) AND NOW()
  AND q.id IN (40)
GROUP BY category,customer


"""
# execute the query and fetch the results
cursor.execute(query)
results = cursor.fetchall()

# create a pandas DataFrame from the query results and transform the Category rows into columns
df = pd.pivot_table(pd.DataFrame(results, columns=['Customer', 'Category', 'Registered']),
                    index='Customer',
                    columns='Category',
                    values='Registered',
                    aggfunc='first',
                    fill_value=0)

# create a copy of the DataFrame to avoid modifying the original
df2 = df.copy()

# add a total row for the number of customers and a total column for the number of categories
df2.loc['Total'] = df2.sum()
df2['Total'] = df2.sum(axis=1)

# export the modified DataFrame to Excel and an HTML table
# create the HTML table without index names
html_table = df2.to_html(index_names=False)

# replace the empty header with "Customer#"
html_table = html_table.replace("<th></th>", "<th>Customer</th>")

# remove the "<th>Category</th>" header from the HTML table
html_table = html_table.replace("<th>Category</th>", "")
print(html_table)



print("SENT")


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
    message['subject'] = 'daily tickets registered report for monitoring queue '
    body = MIMEText(html_table, 'html')
    message.attach(body)
    mail.sendmail("otrs.report@futurenet.in", rcpt, message.as_string())

print("SENT")