import mysql.connector
import xlsxwriter

user = "readuser2"
password = "6FbUDa5VM"
host = "otrs.futurenet.in"
port = 3306
database = 'otrs5'

db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2', password='6FbUDa5VM',
                                  database='otrs5')
cursor = db_conn.cursor()
sql_query = """SELECT  DISTINCT 
(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%%FieldName%%Customer%%' 
and thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%%FieldName%%Customer%%') limit 1) as Customer 
FROM ticket_state ts,users u,ticket_type tt,queue q ,ticket t left join sla s on t.sla_id=s.id 
WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND t.type_id=tt.id AND t.queue_id=q.id 
AND MONTH(t.create_time) = MONTH(CURDATE() - INTERVAL 1 MONTH) 
ORDER BY Customer ASC"""
cursor.execute(sql_query)
results = cursor.fetchall()
#print(results)


