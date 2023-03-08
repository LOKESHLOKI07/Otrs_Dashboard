import mysql.connector
import pandas as pd
pd.set_option('display.max_columns', 10)
import numpy as np

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
sql_query7 = """SELECT distinct
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

data = pd.read_sql(sql_query7,con=db_conn)

stack = data.groupby(["Responsible_user", "type_name"])["type_name"].agg(['count'])
stack_total = stack.groupby("Responsible_user")["count"].sum()


print(stack)
print(stack_total)






