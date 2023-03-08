import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from reportsapp.date3 import cursor


def get_data_from_mysql():
    user = "readuser2"
    password = "6FbUDa5VM"
    host = "otrs.futurenet.in"
    port = 3306
    database = 'otrs5'

    db_conn = mysql.connector.connect(host=host, port=port, user=user, password=password, database=database)
    cursor = db_conn.cursor()

    # Query to get count of tickets by type
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
    sums = sum([i[0] for i in result7])
    print(sums)

    # Query to get ticket data
    sql_query_data = """SELECT distinct t.tn as Ticket_Id,(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,
    t.title as Subject,CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age,
    u.login as Responsible_user,tt.name AS type_name FROM ticket t
    LEFT JOIN ticket_type AS tt ON t.type_id=tt.id 
    LEFT JOIN ticket_state AS ts ON t.ticket_state_id=ts.id 
    LEFT JOIN queue q ON t.queue_id=q.id
    LEFT JOIN users u on  t.user_id=u.id
    WHERE tt.name not in ('-') AND tt.name NOT IN ('junk')
    AND q.name not in ('PRESALES','Misc','Postmaster','ZENTINTL','TEKONCALL','CUSTOMER-ALERTS','ODOO','ODOOHELPDESK')
    AND ts.name  in ('new','ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
    ORDER BY Age DESC"""
    cursor.execute(sql_query_data)
    result_data = cursor.fetchall()
    print(result_data)



    chartdata1 = """SELECT  u.login,
    SUM(CASE WHEN ts.id IN (1, 4, 6, 7, 8, 11, 13, 15, 16, 14, 17, 18) THEN 1 ELSE 0 END) AS 'Work In Process'
    FROM ticket t
    JOIN ticket_state ts ON t.ticket_state_id = ts.id
    JOIN users u ON t.user_id = u.id
    JOIN ticket_type tt ON t.type_id = tt.id
    JOIN queue q ON t.queue_id = q.id
    left JOIN sla s ON s.id = t.sla_id
    WHERE ts.name  in ('new,''ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
    AND  tt.id NOT IN('1')
    AND tt.name NOT IN ('junk')
    AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
    GROUP BY u.login"""
    cursor.execute(chartdata1)
    chartdata01 = cursor.fetchall()
    print(chartdata01)



    idleticket = """SELECT u.login, tt.name,  DATEDIFF(NOW(), t.change_time) AS days_since_change
    FROM ticket t
    JOIN ticket_state ts ON t.ticket_state_id = ts.id
    JOIN users u ON t.user_id = u.id
    JOIN ticket_type tt ON t.type_id = tt.id
    JOIN queue q ON t.queue_id = q.id
    LEFT JOIN sla s ON s.id = t.sla_id
    WHERE ts.name IN ('new','ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
    AND tt.id NOT IN ('1')
    AND tt.name NOT IN ('junk')
    AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
    AND DATEDIFF(NOW(), t.change_time) > 0
    
    GROUP BY t.change_time, u.login, tt.name"""
    cursor.execute(idleticket)
    idleticket01 = cursor.fetchall()

    idle13 = [x[2] for x in idleticket01]
    counts = {
        1: idle13.count(1),
        2: idle13.count(2),
        3: idle13.count(3),
        4: idle13.count(4),
        5: idle13.count(5),
        6: idle13.count(6),
        'greater_than_6': len([x for x in idle13 if x > 6]),
    }

    idlenumbers = list(counts.values())
    print(idlenumbers)

    customerticket = """SELECT(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,
    SUM(CASE WHEN ts.id IN (1, 4, 6, 7, 8, 11, 13, 15, 16, 14, 17, 18) THEN 1 ELSE 0 END) AS 'Work In Process'
    FROM ticket t
    JOIN ticket_state ts ON t.ticket_state_id = ts.id
    JOIN users u ON t.user_id = u.id
    JOIN ticket_type tt ON t.type_id = tt.id
    JOIN queue q ON t.queue_id = q.id
    left JOIN sla s ON s.id = t.sla_id
    WHERE ts.name  in ('new,''ON-HOLD','OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen')
    AND  tt.id NOT IN('1')
    AND tt.name NOT IN ('junk')
    AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
    
    GROUP BY Customer"""
    cursor.execute(customerticket)
    customerticket01 = cursor.fetchall()
    print(customerticket01)
    return result7,result_data,sums,chartdata01,idleticket01,idlenumbers,customerticket01



 # result_data,sums,chartdata01,idleticket01,idlenumbers,customerticket01









