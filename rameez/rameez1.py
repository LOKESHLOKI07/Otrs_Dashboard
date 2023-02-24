import sys
import os
import django

"""sys.path.append('/home/ubuntu/reportproject')
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'reportproject.settings')
django.setup()"""
#from subprocess import run

import smtplib
import pandas as pd
from datetime import date
from datetime import datetime
from datetime import timedelta
import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

today1 = date.today()
today4 = today1.strftime('%d-%m-%Y')
today = today1.strftime('%Y-%m-%d')
today2 = today1 - timedelta(days=1)
today3 = today2.strftime('%Y-%m-%d')
today5 = today2.strftime('%d-%m-%Y')

msg = MIMEMultipart('alternative')
msg['subject'] = 'SLA status as on shift - 14:00:00 - 22:00:00 ' + today3
msg['from'] = 'lokesh.p@futurenet.in'
msg['to'] = 'lokesh.p@futurenet.in'
msg['cc'] = 'lokesh.p@futurenet.in'
msg['bcc'] = 'lokesh.p@futurenet.in'

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'
"""Database Connection"""

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
# query for New Tickets created
sql_query1 = """ SELECT t.tn as Ticket_Id,(select substring(cast(
thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from 
ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS 
Customer, t.title as Subject,u.login as Responsible_user,t.create_time as 
Created_Time,CASE 
WHEN ((SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1)) 
then (SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1) 
ELSE '-'  end AS Response_time FROM ticket t,ticket_state ts,users u,ticket_priority tp,ticket_type tt,queue q,sla s WHERE 
t.ticket_state_id =ts.id AND t.user_id=u.id AND t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND 
t.sla_id=s.id AND t.create_time BETWEEN '""" + today3 + """ 14:00:00' AND '""" + today3 + """ 22:00:00' ORDER BY u.login ASC, t.create_time 
ASC """

cursor.execute(sql_query1)
results1 = cursor.fetchall()
# query for all closed tickets during the shift
sql_query2 = """SELECT t.tn as Ticket_Id, (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' 
in thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id and thi.name like 
'%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from ticket_history thii where 
thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,t.title as Subject,
u.login as Responsible_user, t.create_time as Created_Time, 
CASE 
WHEN ((SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1)) 
then (SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1) 
ELSE '-'  end AS Response_time,
CASE WHEN ts.name IN ( 
'CLOSED SUCCESSFUL','merged', 'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
close-', 'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring( 
cast( thi.name as char( 100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from ticket_history thi 
where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(select max(thii.create_time) from 
ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%' and thi.change_time >='""" + today3 + """ 14:00:00' and thi.change_time <= '""" + today3 + """ 22:00:00' 
and thi.create_time <= '""" + today3 + """ 22:00:00' ) limit 1) as Time_Spent  
FROM ticket t, ticket_state ts, users u, ticket_priority tp, ticket_type tt, queue q WHERE t.ticket_state_id =ts.id 
AND t.user_id=u.id AND t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id IS NOT NULL 
AND CASE WHEN  ts.name IN ('CLOSED SUCCESSFUL', 'merged', 'closed unsuccessful','removed','pending reminder', 
'PENDING AUTO CLOSE','pending auto close-', 'closed with workaround','RESOLVED') then t.change_time >='""" + today3 + """ 14:00:00' and t.change_time <= '""" + today3 + """ 22:00:00' 
and t.create_time <= '""" + today3 + """ 22:00:00' 
END ORDER BY u.login ASC,t.create_time DESC"""
print(sql_query2)
cursor.execute(sql_query2)
results2 = cursor.fetchall()
sql_query3 = """ SELECT t.tn as Ticket_Id,(select substring(cast(
thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from 
ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS 
Customer, t.title as Subject,u.login as Responsible_user,t.create_time as 
Created_Time,CASE 
WHEN ((SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1)) 
then (SELECT min(thi.change_time) FROM ticket_history thi 
WHERE thi.ticket_id = t.id AND thi.history_type_id IN (2,8,11,14) LIMIT 1) 
ELSE '-'  end AS Response_time,(select sum(substring( 
cast( thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32)) as Time_Spent from ticket_history thi 
where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(select max(thii.create_time) from 
ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%' ) limit 1) as Time_Spent 
FROM ticket t,ticket_state ts,users u,ticket_priority tp,ticket_type tt,queue q,sla s WHERE 
t.ticket_state_id =ts.id AND t.user_id=u.id AND t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND 
t.sla_id=s.id and ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') and 
t.create_time < '""" + today3 + """ 14:00:00' ORDER BY u.login ASC, t.create_time ASC """
cursor.execute(sql_query3)
results4 = cursor.fetchall()

"""sql query to dataframe for getting engineer name ,new closed ticket,old closed ticket,Time Spent count """
var5 = pd.DataFrame(results2, columns=['Ticket_id', 'Customer Name', 'Subject', 'Engineer Name', 'Create Time',
                                       'First Response Time',
                                       'Closed Time', 'Time Spent'])

lst5 = []
for x in results2:
    for j in results1:
        if x[0] == j[0]:
            lst5.append(x[0])
            break
    else:
        lst5.append(0)

lst50 = []
for i in lst5:
    if i != 0:
        lst50.append(1)
    else:
        lst50.append(0)
var5['new'] = pd.Series(lst50)
lst51 = []
for j in lst5:
    if j == 0:
        lst51.append(1)
    else:
        lst51.append(0)

var5['old'] = pd.Series(lst51)
df21 = var5[['Engineer Name', 'new']]
df22 = var5[['Engineer Name', 'old']]

import numpy as np

df9 = pd.pivot_table(var5, index='Engineer Name', values=['new', 'old'], aggfunc=np.sum)
df10 = df9.reset_index()
df11 = df9.sum(axis=0)

var6 = var5[["Engineer Name", "Time Spent"]]
var9 = var5[["Engineer Name", "Create Time"]]
df6 = var5.groupby(['Engineer Name'])['Create Time']
df7 = var5.groupby("Engineer Name")["Create Time"].count().reset_index()

var5['Time Spent'] = var5['Time Spent'].replace(np.nan, 0)
var5['Time Spent'] = var5['Time Spent'].astype('int64')
df2 = var5.groupby(['Engineer Name'])['Time Spent'].sum().reset_index()
df2['Time Spent'] = pd.to_datetime(df2["Time Spent"], unit='m').dt.strftime('%H:%M:%S')
result20 = pd.concat([df10, df2['Time Spent']], axis=1)
result = result20.values.tolist()
df3 = df2.values.tolist()

lst55 = []
lst56 = []
for i in df3:
    row1 = "<tr><td>" + str(
        i[0]) + "</td><td align='center'>" + str(i[1]) + "</td></tr>"
    lst56.append(row1)

lst60 = []
lst61 = []
lst62 = []
var = lst56
"""Creating table row for engineer name,new ticket,old ticket,timespent"""
for i in result:
    lst61.append(i[1])
    lst62.append(i[2])
    row1 = "<tr><td>" + str(
        i[0]) + "</td><td style='text-align:center'>" + str(i[1]) + "</td><td style='text-align:center'>" + str(
        i[2]) + "</td><td style='text-align:center'>" + str(i[3]) + "</td></tr>"
    # print(i)
    lst60.append(row1)

lst63 = sum(lst61)
lst64 = sum(lst62)
lst4 = []
"""New ticket closed in between time interval"""
for x in results1:
    # print(x[0])
    for j in results2:
        if x[0] == j[0]:
            # print(x[0])
            lst4.append(x[0])
lst39 = []
for x2 in results2:
    lst39.append(x2[3])
lst40 = set(lst39)
"""Old ticket closed in between time interval """
lst26 = []
for x in results1:
    for j in results2:
        if x[0] != j[0]:
            lst26.append(j[0])

val = list(set(lst26) - set(lst4))
"""finding length of old ticket"""
lst27 = len(val)
""" finding length of new ticket"""
lst31 = len(lst4)
lst8 = []
for y in results1:
    # print(y[5])
    if y[5] == '-':
        lst8.append(y[5])
lst9 = []
lst32 = len(lst8)
for y1 in results2:
    # print(y[5])
    if y1[5] == '-':
        lst8.append(y1[5])

ls2 = []
row = 1
""" table row for query1"""
for i2 in results1:
    for i3 in lst4:
        if i2[0] == i3:
            val = i2[4]
            if i2[5] != '-':
                vall = i2[5]
                vall4 = datetime.strptime(vall, "%Y-%m-%d %H:%M:%S")
                vall15 = vall4.strftime("%d/%m/%Y ,%H:%M:%S")
            else:
                vall15 = '-'
            # print(type(val))
            val2 = val.strftime("%d/%m/%Y ,%H:%M:%S")
            row1 = "<tr bgcolor='#008000' style='color:white'><td style='text-align:center'>" + str(
                row) + "</td><td>" + str(i2[0]) + "</td><td>" + str(i2[1]) + "</td><td>" + str(
                i2[2]) + "</td><td>" + str(
                i2[
                    3]) + "</td><td style='text-align:center'>" + val2 + "</td><td style='text-align:center'>" + vall15 + "</td></tr>"
            ls2.append(row1)
            row += 1
            break
    else:
        val = i2[4]
        if i2[5] != '-':
            vall = i2[5]
            vall4 = datetime.strptime(vall, "%Y-%m-%d %H:%M:%S")
            vall15 = vall4.strftime("%d/%m/%Y ,%H:%M:%S")
        else:
            vall15 = '-'

        val2 = val.strftime("%d/%m/%Y ,%H:%M:%S")
        row3 = "<tr><td style='text-align:center;'>" + str(row) + "</td><td>" + str(i2[0]) + "</td><td>" + str(
            i2[1]) + "</td><td>" + str(i2[2]) + "</td><td>" + str(
            i2[
                3]) + "</td><td style='text-align:center'>" + val2 + "</td><td style='text-align:center'>" + vall15 + "</td></tr>"
        ls2.append(row3)
        row += 1

lst30 = len(ls2)
ls1 = []
row5 = 1
lst34 = []
"""table row for query2"""
for i1 in results2:
    val19 = i1[6]
    lst34.append(val19)

    for i3 in lst4:
        if i1[0] == i3:
            val10 = i1[4]

            if i1[5] != '-':
                val1 = i1[5]
                val4 = datetime.strptime(val1, "%Y-%m-%d %H:%M:%S")
                val15 = val4.strftime("%d/%m/%Y ,%H:%M:%S")
            else:
                val15 = '-'

            val12 = i1[6]
            val14 = val10.strftime("%d/%m/%Y ,%H:%M:%S")
            # val15 = val11.strftime("%d/%m/%Y ,%H:%M:%S")
            val16 = val12.strftime("%d/%m/%Y ,%H:%M:%S")
            row2 = "<tr bgcolor='#008000' style='color:white'><td style='text-align:center'>" + str(
                row5) + "</td><td>" + str(i1[0]) + "</td><td>" + str(i1[1]) + "</td><td>" + str(
                i1[2]) + "</td><td>" + str(i1[3]
                                           ) + "</td><td style='text-align:center'>" + val14 + "</td><td style='text-align:center'>" + val15 + "</td><td style='text-align:center'>" + val16 + "</td><td style='text-align:right'>" + str(
                i1[7]).replace('None', '0') + "</td></tr>"
            ls1.append(row2)
            row5 += 1
            break
    else:
        val10 = i1[4]
        # val11 = i1[5]
        val12 = i1[6]
        if i1[5] != '-':
            val1 = i1[5]
            val4 = datetime.strptime(val1, "%Y-%m-%d %H:%M:%S")
            val15 = val4.strftime("%d/%m/%Y ,%H:%M:%S")
        else:
            val15 = '-'
        # print(val)
        val14 = val10.strftime("%d/%m/%Y ,%H:%M:%S")
        val16 = val12.strftime("%d/%m/%Y ,%H:%M:%S")
        row2 = "<tr><td style='text-align:center'>" + str(row5) + "</td><td>" + str(i1[0]) + "</td><td>" + str(
            i1[1]) + "</td><td>" + str(i1[2]) + "</td><td>" + str(i1[
                                                                      3]) + "</td><td style='text-align:center'>" + val14 + "</td><td style='text-align:center'>" + val15 + "</td><td style='text-align:center'>" + val16 + "</td><td style='text-align:right'>" + str(
            i1[7]).replace('None', '0') + "</td></tr>"
        ls1.append(row2)
        row5 += 1
lst35 = len(lst34)
lst1 = []
row4 = 1

for i2 in results4:
    row3 = "<tr><td style='text-align:center'>" + str(row4) + "</td><td>" + str(i2[0]) + "</td><td>" + str(
        i2[1]) + "</td><td>" + str(i2[2]) + "</td><td>" + str(
        i2[3]) + "</td><td style='text-align:center'>" + str(i2[4]) + "</td><td style='text-align:center'>" + str(
        i2[5]) + "</td><td style='text-align:right'>" + str(i2[6]).replace('None', '0') + "</td></tr>"
    lst1.append(row3)
    row4 += 1

html = """ 

<html><head>
<style>
table, td, th {
  border: 1px solid black;
}

table {
  border-collapse: collapse;
  width: 50%;
}        
</style>
</head>
<body>
<table style="width: 100%">       
<tr>
<center><th colspan="2" bgcolor="yellow" >Summary</th></center>
</tr>
<colgroup>
       <col span="1" style="width: 80%;">
       <col span="1" style="width: 20%;">
</colgroup>       
<tr>
    <td>New tickets registered during the shift</td>
    <td align='center'>""" + str(lst30) + """</td>
</tr>
<tr>
    <td>New Tickets closed within shift</td>
    <td align='center'>""" + str(lst31) + """</td>
</tr>   
<tr>
    <td>First response failed tickets</td>
    <td align='center'>""" + str(lst32) + """</td>
</tr>
<tr>
   <td>Old Tickets Closed within shift</td>
   <td align='center'>""" + str(lst27) + """</td> 
</tr> 
<tr>
    <td>Overall tickets closed during the shift</td>
    <td align='center'>""" + str(lst35) + """</td> 

</tr>
</table>
<br /><br />
<table style="width: 100%">
<tr>
<center><th colspan="4" bgcolor="yellow">Engineer wise closed ticket statistics</th></center>
</tr>
<tr bgcolor="yellow">
    <th>Engineer Name</th>
    <th>New</th>
    <th>Old</th>
    <th>Total Time Spent</th>
</tr>""" + ''.join(lst60) + """ 
<tr bgcolor="grey" style='color:black'>
    <td align='center'>Total Tickets Closed</td>
    <td align='center'>""" + str(lst63) + """</td> 
    <td align='center'>""" + str(lst64) + """</td>
    <td align='center'></td>
</tr></table>

<br /><br />
<table style="width: 100%">
<colgroup>
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 20%;">
       <col span="1" style="width: 35%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 15%;">
       <col span="1" style="width: 10%;">
</colgroup>      
<tr>
<th colspan="7">NEW TICKETS Created during """ + today5 + """ 2:00pm to """ + today5 + """ 10:00pm </th>
</tr>
<tr bgcolor="yellow">
    <th>S.No</th>
    <th>Ticket ID</th>
    <th>Customer Name</th>
    <th>Subject</th>
    <th>Engineer Name</th>
    <th>Create Time</th>
    <th>First Response Time</th>
</tr>""" + ''.join(ls2) + """</table>
<br /><br />
<table style="width: 110%">
<colgroup>
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 20%;">
       <col span="1" style="width: 35%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 15%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 10%;">

</colgroup>      
<tr>
<th colspan="8">OPEN TICKETS Created during """ + str(today5) + """ 2:00pm to """ + str(today5) + """ 10:00pm </th>
</tr>
<tr bgcolor="yellow">
    <th>S.No</th>
    <th>Ticket ID</th>
    <th>Customer Name</th>
    <th>Subject</th>
    <th>Engineer Name</th>
    <th>Create Time</th>
    <th>First Response Time</th>
    <TH>Time Spent </th>
</tr>""" + ''.join(lst1) + """</table>
<br /><br />
<table style="width: 120%">
<colgroup>
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 5%;">
       <col span="1" style="width: 25%;">
       <col span="1" style="width: 35%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 10%;">
       <col span="1" style="width: 10%;">
    </colgroup>
<tr>
<th colspan="9">CLOSED TICKETS during """ + today5 + """ 2:00pm to """ + today5 + """ 10:00pm</th>
</tr>
<tr bgcolor="yellow">

    <th>S.No</th>
    <th>Ticket ID</th>
    <th>Customer Name</th>
    <th>Subject</th>
    <th>Engineer Name</th>
    <th>Create Time</th>
    <th>First Response Time</th>
    <th>Closed Time</th>
    <th>Time Spent</th>
</tr>""" + ''.join(ls1) + """</table>
</html>"""
part2 = MIMEText(html, 'html')
rcpt = msg['cc'].split(",") + [msg['to']]
msg.attach(part2)
mail = smtplib.SMTP('webmail.futurenet.in', 587)
mail.ehlo()
mail.starttls()
mail.login("lokesh.p@futurenet.in", "Classic@123")
mail.sendmail("lokesh.p@futurenet.in", rcpt, msg.as_string())
mail.quit()

"""data = run("rameez1.py",capture_output=True,shell=True)
print(data.stdout)
print(data.stderr)"""