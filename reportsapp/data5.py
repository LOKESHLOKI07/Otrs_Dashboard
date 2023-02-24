# import mysql.connector
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
#
# user = "readuser2"
# password = "6FbUDa5VM"
# host = "otrs.futurenet.in"
# port = 3306
# database = 'otrs5'
#
# db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
#                                   database='otrs5')
# cursor = db_conn.cursor()
# sql_query7 = """SELECT distinct
#                 COUNT( t.tn) as Ticket_Id,
#
#                 tt.name AS type_name
#
# FROM
#                  ticket t,
#                  ticket_state ts,
#                  users u,
#                  ticket_priority tp,
#                  ticket_type tt,
#                  queue q
#                          WHERE
#                  t.ticket_state_id =ts.id
#                 AND t.user_id=u.id
#                  AND t.ticket_priority_id=tp.id
#                  AND t.type_id=tt.id
#                  AND t.queue_id=q.id
#                  AND tt.name not in ('-')
#                  AND q.name not in ('PRESALES','Misc','Postmaster','ZENTINTL','TEKONCALL','CUSTOMER-ALERTS','ODOO','ODOOHELPDESK')
#                   AND ts.name  in ('ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
#                     GROUP BY tt.name
# 		ORDER BY tt.name ASC"""
# cursor.execute(sql_query7)
# result7 = cursor.fetchall()
# print(result7)
#
# len(result7)

# res = [i[0] for i in result7]
# print(res)
# servicerequest =result7[1][0]
# print(servicerequest)

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
sql_query7 = """SELECT COUNT(t.type_id),tt.NAME FROM ticket AS t 
LEFT JOIN ticket_type AS tt ON t.type_id=tt.id 
LEFT JOIN ticket_state AS ts ON t.ticket_state_id=ts.id 
LEFT JOIN queue q ON t.queue_id=q.id
WHERE ts.name  in ('new,''ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
AND  tt.id NOT IN('1')
AND tt.name NOT IN ('junk')
AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
GROUP BY tt.NAME"""
cursor.execute(sql_query7)
result7 = cursor.fetchall()
# print(result7)
sums = sum([i[0] for i in result7])

# res = [i[0] for i in result7]
# print(res)
# servicerequest =result7[1][0]
# print(servicerequest)

sql_query_data = """SELECT distinct
                 t.tn as Ticket_Id,
                (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,
                t.title as Subject,
                 CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age,
                 u.login as Responsible_user,
                         tt.name AS type_name

FROM

                        ticket t
                 LEFT JOIN ticket_type AS tt ON t.type_id=tt.id 
LEFT JOIN ticket_state AS ts ON t.ticket_state_id=ts.id 
LEFT JOIN queue q ON t.queue_id=q.id
LEFT JOIN users u on  t.user_id=u.id
                         WHERE
                 tt.name not in ('-')
                 AND tt.name NOT IN ('junk')
                 AND q.name not in ('PRESALES','Misc','Postmaster','ZENTINTL','TEKONCALL','CUSTOMER-ALERTS','ODOO','ODOOHELPDESK')
                  AND ts.name  in ('new','ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
                    ORDER BY Age DESC"""
cursor.execute(sql_query_data)
result_data = cursor.fetchall()
# print(result7)





