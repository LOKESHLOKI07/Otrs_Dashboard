import io
from django.http import HttpResponse
import mysql.connector
from django.shortcuts import render
from reportsapp.data import results
from reportsapp.data1 import result
# from reportsapp.data5 import result7, result_data,sums,chartdata01,idleticket01,idlenumbers,customerticket01

import xlsxwriter


# def home(requests):
#
#     # servicerequest = {'service1': result7,
#     #                   'length1': len(result7),
#     #                   'full_data': result_data,
#     #                   'sum': sums,
#
#
#                       }




    # return render(requests, 'home.html', servicerequest)
def dashboard(requests):
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
    servicerequest = {
                      'service1': result7,

                      'full_data': result_data,
                      'sum': sums,
                      'chartdata': chartdata01,
                      'idledata': idleticket01,
                      'idlenumbers':idlenumbers,
                      'customerdata': customerticket01,


                      }




    return render(requests, 'dashboard.html', servicerequest)





def customer(request):
    if request.method == "POST":
        if 'Click Download' in request.POST:
            date1 = request.POST['hdate1']
            # print(date1)
            date2 = request.POST['hdate2']
            # print(date2)
            customer_name = request.POST['hname']
            # print(engineer_name)
            if customer_name:
                # print(customer_name)
                var3 = """AND (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in 
                thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id and thi.name like 
                '%%FieldName%%Customer%%' and thi.create_time=(select max(thii.create_time) from ticket_history thii 
                where thii.ticket_id=thi.ticket_id AND thii.name like '%%FieldName%%Customer%%') limit 1) ='""" + customer_name + """' """
                var2 = """ having Customer = '""" + customer_name + """' """
            else:
                var3 = ""
                var2 = ""
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query3 = """SELECT tt.Name,SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
'ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1  ELSE 0 END) AS 'Open', SUM(case WHEN ts.name IN (
'CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
close-','closed with workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ts.name IN ('OPEN',
'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED','pending auto reopen','Waiting for Customer')  THEN 1  ELSE 0 END) AS 'Total', SUM( CASE WHEN 
(TIMESTAMPDIFF(MINUTE,t.create_time,(SELECT min(thi.change_time) FROM ticket_history thi WHERE thi.ticket_id = t.id 
AND thi.history_type_id IN (8) LIMIT 1))/s.first_response_time)  < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN (
'OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED','pending auto reopen','Waiting for Customer')  THEN 1  ELSE 0 END) *100 AS 'ResponseTimeSLA', 
SUM(CASE WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD',
'Waiting for Customer') THEN  NULL WHEN (TIMESTAMPDIFF(MINUTE,t.change_time,NOW())/s.solution_time) < 100 THEN 1 ELSE 
0 END)/SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN 1  ELSE 0 END) * 100 AS 
'ResolutionTimeSLA' FROM  ticket t,ticket_type tt,ticket_state ts, sla s WHERE t.type_id=tt.id  AND  
t.ticket_state_id=ts.id AND t.sla_id = s.id AND tt.name IN ('Incident','Incident Notifications','ServiceRequest',
'Information','Notification') """ + var3 + """ AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' GROUP BY tt.name union SELECT 'Total' as '',SUM(case WHEN ts.name IN 
('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
reopen')  THEN 1  ELSE 0 END) AS 'Open', SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS',
'Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto 
reopen','Waiting for Customer')  THEN 1  ELSE 0 END) AS 'Total', SUM( CASE WHEN (TIMESTAMPDIFF(MINUTE,t.create_time,
(SELECT min(thi.change_time) FROM ticket_history thi WHERE thi.ticket_id = t.id AND thi.history_type_id IN (8) LIMIT 
1))/s.first_response_time)  < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto 
reopen','Waiting for Customer')  THEN 1  ELSE 0 END) *100 AS 'ResponseTimeSLA', SUM(CASE WHEN ts.name IN ('OPEN',
'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') THEN  NULL WHEN (
TIMESTAMPDIFF(MINUTE,t.change_time,NOW())/s.solution_time) < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN ('CLOSED 
SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
'closed with workaround','RESOLVED') THEN 1  ELSE 0 END) * 100 AS 'ResolutionTimeSLA' FROM  ticket t,ticket_type tt,
ticket_state ts, sla s WHERE t.type_id=tt.id  AND  t.ticket_state_id=ts.id AND t.sla_id = s.id AND tt.name IN (
'Incident','Incident Notifications','ServiceRequest','Information','Notification') """ + var3 + """ AND t.create_time 
BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' """
            # print(sql_query3)
            cursor.execute(sql_query3)
            result4 = cursor.fetchall()
            # print(result4)
            sql_query6 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
                        t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
                        Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
                        cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
                        ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
                        select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
                        like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
                        100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
                        thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
                        thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
                        '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
                        position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
                        thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
                        thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
                        from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
                        thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
                        AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
                        ticket_priority tp,ticket_type tt,queue q,sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
                        t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id AND tt.name IN (
                        'ServiceRequest') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
                        """ + var2 + """ ORDER BY t.create_time DESC """
            # print(sql_query6)
            cursor.execute(sql_query6)
            result8 = cursor.fetchall()
            sql_query7 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
                        t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
                        Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
                        cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
                        ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
                        select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
                        like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
                        100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
                        thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
                        thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
                        '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
                        position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
                        thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
                        thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
                        from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
                        thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
                        AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
                        ticket_priority tp,ticket_type tt,queue q,sla s  WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
                        t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
                        'Incident') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' """ + \
                         var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query7)
            result9 = cursor.fetchall()
            # print(result9)
            sql_query8 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
                        t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
                        Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
                        cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
                        ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
                        select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
                        like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
                        100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
                        thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
                        thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
                        '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
                        position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
                        thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
                        thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
                        from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
                        thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
                        AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
                        ticket_priority tp,ticket_type tt,queue q,sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
                        t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
                        'Incident Notifications') AND  t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 
                        23:59:59' """ + var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query8)
            result10 = cursor.fetchall()
            # print(result10)
            sql_query9 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
                        t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
                        Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
                        cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
                        ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
                        select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
                        like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
                        100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
                        thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
                        thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
                        '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
                        position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
                        thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
                        thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
                        from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
                        thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
                        AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
                        ticket_priority tp,ticket_type tt,queue q, sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
                        t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
                        'Information') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
                        """ + var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query9)
            result11 = cursor.fetchall()
            sql_query10 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
                        t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
                        Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
                        cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
                        ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
                        select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
                        like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
                        100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
                        thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
                        thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
                        '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
                        position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
                        thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
                        thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
                        from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
                        thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
                        AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
                        ticket_priority tp,ticket_type tt,queue q, sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
                        t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
                        'Notification') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
                        """ + var3 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query10)
            result15 = cursor.fetchall()

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output)
            worksheet = workbook.add_worksheet()
            row = 0
            col = 0
            bold = workbook.add_format({'bold': True})

            worksheet.write('A1', 'Name', bold)
            worksheet.write('B1', 'Open', bold)
            worksheet.write('C1', 'Closed', bold)
            worksheet.write('D1', 'Total', bold)
            worksheet.write('E1', 'ResponseTimeSLA', bold)
            worksheet.write('F1', 'ResolutionTimeSLA', bold)
            row += 1
            for Name, Open, Closed, Total, ResponseTimeSLA, ResolutionTimeSLA in result4:
                # print(Name)
                worksheet.write(row, col, Name)
                worksheet.write(row, col + 1, Open)
                worksheet.write(row, col + 2, Closed)
                worksheet.write(row, col + 3, Total)
                worksheet.write(row, col + 4, ResponseTimeSLA)
                worksheet.write(row, col + 5, ResolutionTimeSLA)
                row += 1

            row21 = row + 3
            worksheet.write('D' + str(row21), 'Service Request', bold)
            row21 += 2

            worksheet.write('A' + str(row21), 'Customer', bold)
            worksheet.write('B' + str(row21), 'Subject', bold)
            worksheet.write('C' + str(row21), 'Sender', bold)
            worksheet.write('D' + str(row21), 'Responsible_user', bold)
            worksheet.write('E' + str(row21), 'Status_Name', bold)
            worksheet.write('F' + str(row21), 'Created_Date', bold)
            worksheet.write('G' + str(row21), 'Created_time', bold)
            worksheet.write('H' + str(row21), 'Closed_Date', bold)
            worksheet.write('I' + str(row21), 'Closed_time', bold)
            worksheet.write('J' + str(row21), 'Time_Spent', bold)
            worksheet.write('K' + str(row21), 'SLA', bold)
            worksheet.write('L' + str(row21), 'category', bold)
            worksheet.write('M' + str(row21), 'sources', bold)
            worksheet.write('N' + str(row21), 'Customer', bold)

            # col1 = 0
            # print(row1)

            for Ticket_ID, Subject, Sender, Responsible_user, Status_Name, Created_Time, Closed_time, Time_Spent, SLA, category, sources, Customer in result8:
                dates, times = str(Created_Time).split(' ')
                times1 = str(Closed_time).replace('None', ' ')
                # print(times1)
                dates2, times2 = times1.split(' ')

                worksheet.write(row21, col, Ticket_ID)
                worksheet.write(row21, col + 1, Subject)
                worksheet.write(row21, col + 2, Sender)
                worksheet.write(row21, col + 3, Responsible_user)
                worksheet.write(row21, col + 4, Status_Name)
                worksheet.write(row21, col + 5, dates)
                worksheet.write(row21, col + 6, times)
                worksheet.write(row21, col + 7, dates2)
                worksheet.write(row21, col + 8, times2)
                worksheet.write(row21, col + 9, Time_Spent)
                worksheet.write(row21, col + 10, SLA)
                worksheet.write(row21, col + 11, category)
                worksheet.write(row21, col + 12, sources)
                worksheet.write(row21, col + 13, Customer)
                row21 += 1

            row22 = row21 + 3
            worksheet.write('D' + str(row22), 'Incident', bold)
            row22 += 2
            worksheet.write('A' + str(row22), 'Ticket_ID', bold)
            worksheet.write('B' + str(row22), 'Subject', bold)
            worksheet.write('C' + str(row22), 'Sender', bold)
            worksheet.write('D' + str(row22), 'Responsible_user', bold)
            worksheet.write('E' + str(row22), 'Status_Name', bold)
            worksheet.write('F' + str(row22), 'Created_Date', bold)
            worksheet.write('G' + str(row22), 'Created_time', bold)
            worksheet.write('H' + str(row22), 'Closed_Date', bold)
            worksheet.write('I' + str(row22), 'Closed_time', bold)
            worksheet.write('J' + str(row22), 'Time_Spent', bold)
            worksheet.write('K' + str(row22), 'SLA', bold)
            worksheet.write('L' + str(row22), 'category', bold)
            worksheet.write('M' + str(row22), 'sources', bold)
            worksheet.write('N' + str(row22), 'Customer', bold)

            for Ticket_ID1, Subject1, Sender1, Responsible_user1, Status_Name1, Created_Time1, Closed_time1, Time_Spend1, SLA1, category1, sources1, Customer1 in result9:
                dates1, times2 = str(Created_Time1).split(' ')
                times1 = str(Closed_time1).replace('None', ' ')
                # print(times1)
                dates3, times3 = times1.split(' ')
                worksheet.write(row22, col, Ticket_ID1)
                worksheet.write(row22, col + 1, Subject1)
                worksheet.write(row22, col + 2, Sender1)
                worksheet.write(row22, col + 3, Responsible_user1)
                worksheet.write(row22, col + 4, Status_Name1)
                worksheet.write(row22, col + 5, dates1)
                worksheet.write(row22, col + 6, times2)
                worksheet.write(row22, col + 7, dates3)
                worksheet.write(row22, col + 8, times3)
                worksheet.write(row22, col + 9, Time_Spend1)
                worksheet.write(row22, col + 10, SLA1)
                worksheet.write(row22, col + 11, category1)
                worksheet.write(row22, col + 12, sources1)
                worksheet.write(row22, col + 13, Customer1)
                row22 += 1

            row23 = row22 + 3
            worksheet.write('D' + str(row23), 'Incident Notification', bold)
            row23 += 2
            worksheet.write('A' + str(row23), 'Ticket_ID', bold)
            worksheet.write('B' + str(row23), 'Subject', bold)
            worksheet.write('C' + str(row23), 'Sender', bold)
            worksheet.write('D' + str(row23), 'Responsible_user', bold)
            worksheet.write('E' + str(row23), 'Status_Name', bold)
            worksheet.write('F' + str(row23), 'Created_Date', bold)
            worksheet.write('G' + str(row23), 'Created_time', bold)
            worksheet.write('H' + str(row23), 'Closed_Date', bold)
            worksheet.write('I' + str(row23), 'Closed_time', bold)
            worksheet.write('J' + str(row23), 'Time_Spent', bold)
            worksheet.write('K' + str(row23), 'SLA', bold)
            worksheet.write('L' + str(row23), 'category', bold)
            worksheet.write('M' + str(row23), 'sources', bold)
            worksheet.write('N' + str(row23), 'Customer', bold)

            for Ticket_ID2, Subject2, Sender2, Responsible_user2, Status_Name2, Created_Time2, Closed_time2, Time_Spend2, SLA2, category2, sources2, Customer2 in result10:
                dates2, times4 = str(Created_Time2).split(' ')
                times1 = str(Closed_time2).replace('None', ' ')
                # print(times1)
                dates4, times5 = times1.split(' ')
                worksheet.write(row23, col, Ticket_ID2)
                worksheet.write(row23, col + 1, Subject2)
                worksheet.write(row23, col + 2, Sender2)
                worksheet.write(row23, col + 3, Responsible_user2)
                worksheet.write(row23, col + 4, Status_Name2)
                worksheet.write(row23, col + 5, dates2)
                worksheet.write(row23, col + 6, times4)
                worksheet.write(row23, col + 7, dates4)
                worksheet.write(row23, col + 8, times5)
                worksheet.write(row23, col + 9, Time_Spend2)
                worksheet.write(row23, col + 10, SLA2)
                worksheet.write(row23, col + 11, category2)
                worksheet.write(row23, col + 12, sources2)
                worksheet.write(row23, col + 13, Customer2)
                row23 += 1

            row24 = row23 + 3
            worksheet.write('D' + str(row24), 'Information', bold)
            row24 += 2
            worksheet.write('A' + str(row24), 'Ticket_ID', bold)
            worksheet.write('B' + str(row24), 'Subject', bold)
            worksheet.write('C' + str(row24), 'Sender', bold)
            worksheet.write('D' + str(row24), 'Responsible_user', bold)
            worksheet.write('E' + str(row24), 'Status_Name', bold)
            worksheet.write('F' + str(row24), 'Created_Date', bold)
            worksheet.write('G' + str(row24), 'Created_time', bold)
            worksheet.write('H' + str(row24), 'Closed_Date', bold)
            worksheet.write('I' + str(row24), 'Closed_time', bold)
            worksheet.write('J' + str(row24), 'Time_Spent', bold)
            worksheet.write('K' + str(row24), 'SLA', bold)
            worksheet.write('L' + str(row24), 'category', bold)
            worksheet.write('M' + str(row24), 'sources', bold)
            worksheet.write('N' + str(row24), 'Customer', bold)

            for Ticket_ID3, Subject3, Sender3, Responsible_user3, Status_Name3, Created_Time3, Closed_time3, Time_Spend3, SLA3, category3, sources3, Customer3 in result11:
                dates5, times6 = str(Created_Time3).split(' ')
                times1 = str(Closed_time3).replace('None', ' ')
                # print(times1)
                dates6, times7 = times1.split(' ')
                worksheet.write(row24, col, Ticket_ID3)
                worksheet.write(row24, col + 1, Subject3)
                worksheet.write(row24, col + 2, Sender3)
                worksheet.write(row24, col + 3, Responsible_user3)
                worksheet.write(row24, col + 4, Status_Name3)
                worksheet.write(row24, col + 5, dates5)
                worksheet.write(row24, col + 6, times6)
                worksheet.write(row24, col + 7, dates6)
                worksheet.write(row24, col + 8, times7)
                worksheet.write(row24, col + 9, Time_Spend3)
                worksheet.write(row24, col + 10, SLA3)
                worksheet.write(row24, col + 11, category3)
                worksheet.write(row24, col + 12, sources3)
                worksheet.write(row24, col + 13, Customer3)
                row24 += 1

            row25 = row24 + 3
            worksheet.write('D' + str(row25), 'Notification', bold)
            row25 += 2
            worksheet.write('A' + str(row25), 'Ticket_ID', bold)
            worksheet.write('B' + str(row25), 'Subject', bold)
            worksheet.write('C' + str(row25), 'Sender', bold)
            worksheet.write('D' + str(row25), 'Responsible_user', bold)
            worksheet.write('E' + str(row25), 'Status_Name', bold)
            worksheet.write('F' + str(row25), 'Created_Date', bold)
            worksheet.write('G' + str(row25), 'Created_time', bold)
            worksheet.write('H' + str(row25), 'Closed_Date', bold)
            worksheet.write('I' + str(row25), 'Closed_time', bold)
            worksheet.write('J' + str(row25), 'Time_Spent', bold)
            worksheet.write('K' + str(row25), 'SLA', bold)
            worksheet.write('L' + str(row25), 'category', bold)
            worksheet.write('M' + str(row25), 'sources', bold)
            worksheet.write('N' + str(row25), 'Customer', bold)

            for Ticket_ID4, Subject4, Sender4, Responsible_user4, Status_Name4, Created_Time4, Closed_time4, Time_Spend4, SLA4, category4, sources4, Customer4 in result15:
                dates7, times8 = str(Created_Time4).split(' ')
                times1 = str(Closed_time4).replace('None', ' ')
                # print(times1)
                dates8, times9 = times1.split(' ')
                worksheet.write(row25, col, Ticket_ID4)
                worksheet.write(row25, col + 1, Subject4)
                worksheet.write(row25, col + 2, Sender4)
                worksheet.write(row25, col + 3, Responsible_user4)
                worksheet.write(row25, col + 4, Status_Name4)
                worksheet.write(row25, col + 5, dates7)
                worksheet.write(row25, col + 6, times8)
                worksheet.write(row25, col + 7, dates8)
                worksheet.write(row25, col + 8, times9)
                worksheet.write(row25, col + 9, Time_Spend4)
                worksheet.write(row25, col + 10, SLA4)
                worksheet.write(row25, col + 11, category4)
                worksheet.write(row25, col + 12, sources4)
                worksheet.write(row25, col + 13, Customer4)
                row25 += 1
            workbook.close()
            output.seek(0)
            filename = 'Customer Report.xlsx'
            response = HttpResponse(
                output,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=%s' % filename
            return response

        else:

            date1 = request.POST['date1']
            date2 = request.POST['date2']
            customer_name = request.POST['customer_name']
            if customer_name:

                var3 = """ AND (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in 
                thi.name)-31) as Customer from ticket_history thi where thi.ticket_id=t.id and thi.name like 
                '%%FieldName%%Customer%%' and thi.create_time=(select max(thii.create_time) from ticket_history thii 
                where thii.ticket_id=thi.ticket_id AND thii.name like '%%FieldName%%Customer%%') limit 1) ='""" + customer_name + """'  """
                var2 = """ having Customer = '""" + customer_name + """' """
            else:
                var3 = ""
                var2 = ""
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query3 = """ SELECT tt.Name,SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
'ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1  ELSE 0 END) AS 'Open', SUM(case WHEN ts.name IN (
'CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
close-','closed with workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ts.name IN ('OPEN',
'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED','pending auto reopen','Waiting for Customer')  THEN 1  ELSE 0 END) AS 'Total', SUM( CASE WHEN 
(TIMESTAMPDIFF(MINUTE,t.create_time,(SELECT min(thi.change_time) FROM ticket_history thi WHERE thi.ticket_id = t.id 
AND thi.history_type_id IN (8) LIMIT 1))/s.first_response_time)  < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN (
'OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED','pending auto reopen','Waiting for Customer')  THEN 1  ELSE 0 END) *100 AS 'ResponseTimeSLA', 
SUM(CASE WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD',
'Waiting for Customer') THEN  NULL WHEN (TIMESTAMPDIFF(MINUTE,t.change_time,NOW())/s.solution_time) < 100 THEN 1 ELSE 
0 END)/SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN 1  ELSE 0 END) * 100 AS 
'ResolutionTimeSLA' FROM  ticket t,ticket_type tt,ticket_state ts, sla s WHERE t.type_id=tt.id  AND  
t.ticket_state_id=ts.id AND t.sla_id = s.id AND tt.name IN ('Incident','Incident Notifications','ServiceRequest',
'Information','Notification') """ + var3 + """ AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' GROUP BY tt.name union SELECT 'Total' as '',SUM(case WHEN ts.name IN 
('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
reopen')  THEN 1  ELSE 0 END) AS 'Open', SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS',
'Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto 
reopen','Waiting for Customer')  THEN 1  ELSE 0 END) AS 'Total', SUM( CASE WHEN (TIMESTAMPDIFF(MINUTE,t.create_time,
(SELECT min(thi.change_time) FROM ticket_history thi WHERE thi.ticket_id = t.id AND thi.history_type_id IN (8) LIMIT 
1))/s.first_response_time)  < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto 
reopen','Waiting for Customer')  THEN 1  ELSE 0 END) *100 AS 'ResponseTimeSLA', SUM(CASE WHEN ts.name IN ('OPEN',
'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') THEN  NULL WHEN (
TIMESTAMPDIFF(MINUTE,t.change_time,NOW())/s.solution_time) < 100 THEN 1 ELSE 0 END)/SUM(case WHEN ts.name IN ('CLOSED 
SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
'closed with workaround','RESOLVED') THEN 1  ELSE 0 END) * 100 AS 'ResolutionTimeSLA' FROM  ticket t,ticket_type tt,
ticket_state ts, sla s WHERE t.type_id=tt.id  AND  t.ticket_state_id=ts.id AND t.sla_id = s.id AND tt.name IN (
'Incident','Incident Notifications','ServiceRequest','Information','Notification') """ + var3 + """ AND t.create_time 
BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' """

            # print(sql_query3)
            cursor.execute(sql_query3)
            result4 = cursor.fetchall()
            # print(result4)
            sql_query6 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
            t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
            Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
            cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
            ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
            select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
            like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
            100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
            thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
            thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
            '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
            position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
            thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
            thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
            (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
            from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
            thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
            AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
            ticket_priority tp,ticket_type tt,queue q,sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
            t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id AND tt.name IN (
            'ServiceRequest') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
            """ + var2 + """ ORDER BY t.create_time DESC """

            # print(sql_query6)
            cursor.execute(sql_query6)
            result8 = cursor.fetchall()
            sql_query7 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
            t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
            Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
            cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
            ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
            select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
            like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
            100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
            thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
            thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
            '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
            position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
            thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
            thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
            (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
            from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
            thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
            AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
            ticket_priority tp,ticket_type tt,queue q,sla s  WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
            t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
            'Incident') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' """ + \
                         var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query7)
            result9 = cursor.fetchall()
            # print(result9)
            sql_query8 = """ SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
            t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
            Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
            cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
            ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
            select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
            like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
            100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
            thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
            thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
            '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
            position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
            thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
            thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
            (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
            from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
            thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
            AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
            ticket_priority tp,ticket_type tt,queue q,sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
            t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
            'Incident Notifications') AND  t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 
            23:59:59' """ + var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query8)
            result10 = cursor.fetchall()
            # print(result10)
            sql_query9 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
            t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
            Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
            cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
            ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
            select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
            like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
            100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
            thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
            thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
            '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
            position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
            thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
            thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
            (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
            from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
            thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
            AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
            ticket_priority tp,ticket_type tt,queue q, sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
            t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
            'Information') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
            """ + var2 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query9)
            result11 = cursor.fetchall()
            sql_query10 = """SELECT t.tn as Ticket_Id, t.title as Subject,(case when t.customer_id IS NOT NULL then 
            t.customer_id ELSE 'support@futurenet.in' END)  AS Sender,u.login as Responsible_user,ts.name as 
            Status_Name,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time, (select substring(
            cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent from 
            ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(
            select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name 
            like '%TimeSpent%%%%') limit 1) as Time_Spent,s.name as SLA, (select substring(cast(thi.name as char(
            100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer from ticket_history thi where 
            thi.ticket_id=t.id AND thi.name like '%FieldName%Category%' and thi.create_time=(select max(
            thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id AND thii.name like 
            '%FieldName%Category%') LIMIT 1 ) AS category , (select substring(cast(thi.name as char(100)),'29',
            position('%%OldValue%%' in thi.name)-29) as Customer from ticket_history thi where thi.ticket_id=t.id AND 
            thi.name like '%FieldName%Source%' and thi.create_time=(select max(thii.create_time) from ticket_history 
            thii where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') LIMIT 1 ) AS sources , 
            (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer 
            from ticket_history thi where thi.ticket_id=t.id AND thi.name like '%FieldName%Customer%' and 
            thi.create_time=(select max(thii.create_time) from ticket_history thii where thii.ticket_id=thi.ticket_id 
            AND thii.name like '%FieldName%Customer%') LIMIT 1 ) AS Customer FROM ticket t,ticket_state ts,users u,
            ticket_priority tp,ticket_type tt,queue q, sla s WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND 
            t.ticket_priority_id=tp.id AND t.type_id=tt.id AND t.queue_id=q.id AND t.sla_id = s.id  AND tt.name IN (
            'Notification') AND t.create_time BETWEEN '""" + date1 + """ 00:00:00' AND '""" + date2 + """ 23:59:59' 
            """ + var3 + """ ORDER BY t.create_time DESC """
            cursor.execute(sql_query10)
            result15 = cursor.fetchall()
            context3 = {
                'customer': result4, 'customer1': result8, 'customer2': result9, 'customer3': result10,
                'customer4': result11, 'customer5': result15, 'hname': customer_name, 'hdate1': date1, 'hdate2': date2
            }
            # print(context3)
            # print(len(context2))
            return render(request, 'base1.html', context3)
    else:
        context1 = {
            'members': results
        }
        return render(request, 'customerreport.html', context1)


def engineer(request):
    if request.method == "POST":
        if 'Click Download' in request.POST:
            date1 = request.POST['hdate1']
            # print(date1)
            date2 = request.POST['hdate2']
            # print(date2)
            engineer_name = request.POST['hname']
            # print(engineer_name)
            if engineer_name:

                var1 = """ u.login = '""" + str(engineer_name) + """' and """
            else:
                var1 = ""
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query1 = """SELECT u.login as Engineer_Name, SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS',
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') THEN 1 ELSE 0 END) AS 'Open', 
            SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto reopen') THEN 1 
            ELSE 0 END ) AS 'Closed' , COUNT(*) as Total FROM  ticket t, ticket_state ts, users u, ticket_priority tp, 
            ticket_type tt, queue q WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND t.ticket_priority_id=tp.id AND 
            t.type_id=tt.id AND t.queue_id=q.id AND CASE WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') then  t.create_time <= '""" + str(date2) + """ 23:59:59' 
            WHEN  ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto reopen') then  
            t.create_time between '""" + str(date1) + """' and '""" + str(date2) + """ 23:59:59'  END and """ + var1 + """
            t.sla_id IS NOT NULL GROUP BY u.login ASC """
            # print(sql_query1)
            cursor.execute(sql_query1)
            result2 = cursor.fetchall()
            sql_query12 = """SELECT distinct
            t.tn as Ticket_Id,t.title as Subject,(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
t.create_time as Created_time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time,
ts.name as state,sl.name AS sla,tt.name Type_Name,q.name as queue ,se.name AS service,
(select substring(cast(thi.name as char(100)),'29',position('%%OldValue%%' in thi.name)-29) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Source%' 
and thi.create_time=(select max(thii.create_time) from ticket_history thii 
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') limit 1) as Source, 
(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi 
where thi.ticket_id=t.id and thi.name like '%FieldName%Customer%' 
and thi.create_time=(select max(thii.create_time) from ticket_history thii 
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer,
CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age,
u.login as Responsible_user,
(select substring(cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent
from ticket_history thi
where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%') limit 1) as Time_Spent
FROM 
ticket_state ts,users u,ticket_priority tp,ticket t 
LEFT JOIN ticket_type tt ON t.type_id = tt.id 
LEFT JOIN sla sl ON t.sla_id = sl.id 
LEFT JOIN service se ON t.service_id = se.id 
LEFT JOIN queue q ON t.queue_id=q.id 
WHERE """ + str(var1) + """ t.ticket_state_id =ts.id AND t.user_id=u.id AND tt.name NOT IN ('junk') 
AND ts.name NOT IN ('merged') AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK')  
and t.create_time BETWEEN '""" + str(date1) + """' and '""" + str(date2) + """ 23:59:59' ORDER BY t.create_time DESC """
            # print(sql_query12)
            cursor.execute(sql_query12)
            result12 = cursor.fetchall()

            # print(result2)
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output)
            worksheet = workbook.add_worksheet()
            row = 0
            col = 0
            # lst=['Responsible_user', 'Open', 'Closed', 'Total']
            # worksheet.set_column(lst)
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'Responsible_user', bold)
            worksheet.write('B1', 'Open', bold)
            worksheet.write('C1', 'Closed', bold)
            worksheet.write('D1', 'Total', bold)
            row += 1
            for Responsible_user, Open, Closed, Total in result2:
                worksheet.write(row, col, Responsible_user)
                worksheet.write(row, col + 1, Open)
                worksheet.write(row, col + 2, Closed)
                worksheet.write(row, col + 3, Total)
                row += 1
            row12 = row + 3
            # row14 = row12 + 3

            worksheet.write('A' + str(row12), 'Ticket#', bold)
            worksheet.write('B' + str(row12), 'Title', bold)
            worksheet.write('C' + str(row12), 'Category', bold)
            worksheet.write('D' + str(row12), 'Created_date', bold)
            worksheet.write('E' + str(row12), 'Created_time', bold)
            worksheet.write('F' + str(row12), 'Closed_date', bold)
            worksheet.write('G' + str(row12), 'Closed_time', bold)
            worksheet.write('H' + str(row12), 'State', bold)
            worksheet.write('I' + str(row12), 'SLA', bold)
            worksheet.write('J' + str(row12), 'Type', bold)
            worksheet.write('K' + str(row12), 'Queue', bold)
            worksheet.write('L' + str(row12), 'Service', bold)
            worksheet.write('M' + str(row12), 'Ticket_Source', bold)
            worksheet.write('N' + str(row12), 'Customer Name', bold)
            worksheet.write('O' + str(row12), 'No.of Articles', bold)
            worksheet.write('P' + str(row12), 'Agent/owner', bold)
            worksheet.write('Q' + str(row12), 'Time_Spent', bold)

            for Ticket_ID, Title, Category, Created_time, Closed_time, State, SLA, Type, Queue, Service, Ticket_source, Customer, Age, Engineer_Name, Time_Spent in result12:
                dates, times = str(Created_time).split(' ')
                times1 = str(Closed_time).replace('None', ' ')
                # print(times1)
                dates2, times2 = times1.split(' ')

                # print(dates2, times2)
                worksheet.write(row12, col, Ticket_ID)
                worksheet.write(row12, col + 1, Title)
                worksheet.write(row12, col + 2, Category)
                worksheet.write(row12, col + 3, dates)
                worksheet.write(row12, col + 4, times)
                worksheet.write(row12, col + 5, dates2)
                worksheet.write(row12, col + 6, times2)
                worksheet.write(row12, col + 7, State)
                worksheet.write(row12, col + 8, SLA)
                worksheet.write(row12, col + 9, Type)
                worksheet.write(row12, col + 10, Queue)
                worksheet.write(row12, col + 11, Service)
                worksheet.write(row12, col + 12, Ticket_source)
                worksheet.write(row12, col + 13, Customer)
                worksheet.write(row12, col + 14, Age)
                worksheet.write(row12, col + 15, Engineer_Name)
                worksheet.write(row12, col + 16, Time_Spent)

                row12 += 1

            workbook.close()
            output.seek(0)
            filename = 'Engineer Report.xlsx'
            response = HttpResponse(
                output,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=%s' % filename
            return response


        else:
            date1 = request.POST['date1']
            date2 = request.POST['date2']
            engineer_name = request.POST['engineer_name']

            if engineer_name:

                var1 = """ u.login= '""" + str(engineer_name) + """' and """
            else:
                var1 = ""
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query1 = """SELECT u.login as Responsible_user, SUM(case WHEN ts.name IN ('OPEN','WORK IN PROGRESS',
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') THEN 1 ELSE 0 END) AS 'Open', 
            SUM(case WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto reopen') THEN 1 
            ELSE 0 END ) AS 'Closed' , COUNT(*) as Total FROM  ticket t, ticket_state ts, users u, ticket_priority tp, 
            ticket_type tt, queue q WHERE t.ticket_state_id =ts.id AND t.user_id=u.id AND t.ticket_priority_id=tp.id AND 
            t.type_id=tt.id AND t.queue_id=q.id AND CASE WHEN ts.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer') then  t.create_time <= '""" + date2 + """ 23:59:59' 
            WHEN  ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED','pending auto reopen') then  
            t.create_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59'  END and """ + var1 + """
            t.sla_id IS NOT NULL GROUP BY u.login ASC """
            # print(sql_query1)
            cursor.execute(sql_query1)
            result3 = cursor.fetchall()
            sql_query12 = """SELECT distinct
t.tn as Ticket_Id,
u.login as Responsible_user,
tt.name Type_Name ,
t.title as Subject,
(select substring(cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent
from ticket_history thi
where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%' and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%') limit 1) as Time_Spent
,(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi 
where thi.ticket_id=t.id and thi.name like '%FieldName%Customer%' 
and thi.create_time=(select max(thii.create_time) from ticket_history thii 
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer
,t.create_time as Created_time,
CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time,                
CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age,
q.name,sl.name AS sla,se.name AS service,ts.name,
(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
and thi.create_time=(select max(thii.create_time) from ticket_history thii
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
(select substring(cast(thi.name as char(100)),'29',position('%%OldValue%%' in thi.name)-29) as Customer
from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Source%' 
and thi.create_time=(select max(thii.create_time) from ticket_history thii 
where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') limit 1) as Source 
FROM 
ticket_state ts,users u,ticket_priority tp,ticket t 
LEFT JOIN ticket_type tt ON t.type_id = tt.id 
LEFT JOIN sla sl ON t.sla_id = sl.id 
LEFT JOIN service se ON t.service_id = se.id 
LEFT JOIN queue q ON t.queue_id=q.id 
WHERE """ + var1 + """ t.ticket_state_id =ts.id AND t.user_id=u.id AND tt.name NOT IN ('junk') 
AND ts.name NOT IN ('merged') AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK') 
and t.create_time BETWEEN '""" + date1 + """' AND '""" + date2 + """ 23:59:59' ORDER BY t.create_time DESC """
            # print(sql_query12)
            cursor.execute(sql_query12)
            result12 = cursor.fetchall()

            context2 = {

                'engineer': result3, 'engineer1': result12, 'hname': engineer_name, 'hdate1': date1, 'hdate2': date2
            }
            return render(request, 'base.html', context2)



    else:
        context = {
            'member': result
        }
        return render(request, 'engineerreport.html', context)


def fullsummary(request):
    if request.method == "POST":
        if 'Click Download' in request.POST:
            date1 = request.POST['date1']
            # print(date1)
            date2 = request.POST['date2']
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query21 = """SELECT 'COMPLETED TICKETS' AS 'TICKETS', SUM(case WHEN ticket_state.name IN ('CLOSED 
                        SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
                        close-','closed with workaround','RESOLVED') THEN 1  END) AS 'COMPLETED TICKETS' FROM ticket_state,ticket, 
                        queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS NOT NULL 
                        and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor', 
                        'ON-HOLD','Waiting for Customer','pending auto reopen') then  ticket.change_time <= '""" + date2 + """ 23:59:59' when 
                        ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder', 
                        'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
                        between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end UNION SELECT 'OPENED TICKETS' AS 'TICKETS', SUM(case WHEN 
                        ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD',
                        'Waiting for Customer','pending auto reopen')  THEN 1 ELSE 0  END) AS 'OPEN TICKETS' FROM ticket_state,
                        ticket,queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS 
                        NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for 
                        Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then ticket.create_time <= '""" + date2 + """ 23:59:59' 
                        when ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
                        'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then ticket.change_time 
                        between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end union 
                        SELECT 'TOTAL NO. OF TICKETS' AS 'TICKETS',SUM(case WHEN ticket_state.name IN ('CLOSED 
                        SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
                        close-','closed with workaround','RESOLVED','pending auto reopen','Waiting for Customer','OPEN',
                        'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD')THEN 1  END )AS 'TOTAL' FROM 
                        ticket_state,ticket,queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND 
                        ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
                        Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then  
                        ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end  """
            cursor.execute(sql_query21)
            result15 = cursor.fetchall()
            sql_query23 = """SELECT ticket_type.Name,SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 
            1 ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful', 'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ticket_state.name IN (
            'OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL',
            'merged','closed unsuccessful','removed', 'pending reminder', 'PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED', 'pending auto reopen', 'Waiting for Customer') THEN 1 ELSE 0 END) AS 
            'Total' FROM ticket,ticket_type, ticket_state WHERE ticket.type_id=ticket_type.id AND 
            ticket.ticket_state_id=ticket_state.id AND ticket.sla_id IS NOT NULL /*	AND ticket.change_time >= '""" + \
                          date1 + """' AND ticket.change_time <= '""" + date2 + """23 :59:59' */ and case when 
            ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for 
            Vendor','ON-HOLD', 'Waiting for Customer','pending auto reopen') then  ticket.create_time 
            <= '""" + date2 + """ 23:59:59' when 
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """23:59:59' end GROUP BY ticket_type.id union 
            SELECT 'Total' as 'Name', SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
            reopen')  THEN 1 ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED 
            SUCCESSFUL','merged','closed unsuccessful', 'removed','pending reminder','PENDING AUTO 
            CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN 1  ELSE 0 END) AS 
            'Closed', count(ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval', 
            'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed', 
            'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround',
            'RESOLVED', 'pending auto reopen','Waiting for Customer')) as 'Total' FROM  ticket,
            ticket_type,ticket_state WHERE ticket.type_id=ticket_type.id AND  
            ticket.ticket_state_id=ticket_state.id AND ticket.sla_id IS NOT NULL /* AND 
            ticket.change_time >= '""" + date1 + """' AND ticket.change_time <= '""" + date2 + """23:59
            :59'*/ and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval',
            'Waiting for Vendor','ON-HOLD', 'Waiting for Customer','pending auto reopen') then  
            ticket.create_time <= '""" + date2 + """23:59:59' when ticket_state.name IN ('CLOSED 
            SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder', 'PENDING AUTO 
            CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query23)
            result17 = cursor.fetchall()
            sql_query25 = """SELECT  queue.Name AS Queue, SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS',
                        'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
                        ELSE 0  END) AS 'Open' , SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED','pending auto reopen','Waiting for Customer')) AS 'Total' FROM ticket_state,ticket,
                        queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS NOT NULL 
                        and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
                        'ON-HOLD','Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when  
                        ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
                        'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
                        between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end GROUP BY queue.Name union 
                        SELECT 'Total' AS 'Queue', SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
                        'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
                        ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful', 'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN ('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval', 'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED', 'pending auto reopen','Waiting for Customer')) AS 'Total' FROM ticket,ticket_type,
                        ticket_state WHERE ticket.type_id=ticket_type.id AND  ticket.ticket_state_id=ticket_state.id AND
                        ticket.sla_id IS NOT NULL /* AND ticket.change_time >= '""" + date1 + """' AND ticket.change_time <= '""" + date2 + """ 23:59:59'*/ 
                        and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
                        'ON-HOLD', 'Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:50' when 
                        ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder', 
                        'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
                        between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query25)
            result19 = cursor.fetchall()
            sql_query27 = """SELECT  users.Login , SUM(CASE WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
                        'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
                        ELSE 0  END) AS 'Open', SUM(CASE WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged', 
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') THEN 1 ELSE 0 END) AS 'Closed', SUM(CASE WHEN ticket_state.name IN ('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder', 'PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED','pending auto reopen', 'Waiting for Customer')  THEN 1 ELSE 0  END) AS 'Total' FROM 
                        ticket, users ,ticket_state WHERE ticket.user_id=users.id AND ticket.ticket_state_id=ticket_state.id AND 
                        ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
                        Approval','Waiting for Vendor', 'ON-HOLD','Waiting for Customer','pending auto reopen') then  
                        ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' 
                        end GROUP BY users.Login union 
                        SELECT 'Total' AS 'Login',SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
                        'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
                        ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful', 'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN ('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval', 'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED', 'pending auto reopen','Waiting for Customer')) AS 'Total' FROM    ticket,ticket_type,
                        ticket_state WHERE ticket.type_id=ticket_type.id AND  ticket.ticket_state_id=ticket_state.id AND 
                        ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
                        Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then 
                        ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL', 'merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE', 'pending auto close-','closed with 
                        workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end"""
            cursor.execute(sql_query27)
            result21 = cursor.fetchall()
            sql_query29 = """SELECT  service.name AS Services , SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  
                        END) AS 'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 
                        'Incident::Major', SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', 
                        SUM(case WHEN ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem',SUM(case WHEN ticket_type.name 
                        ='Report'  THEN 1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  
                        END) AS 'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
                        ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
                        ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
                        'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
                        'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'TOTAL' FROM ticket, service,
                        ticket_type,ticket_state WHERE  ticket.service_id=service.id and ticket.ticket_state_id=ticket_state.id AND 
                        ticket.type_id=ticket_type.id and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
                        Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then  
                        ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') then  ticket.change_time between '""" + date1 + """' and '""" + date2 + """23:59:59' end AND ticket.sla_id IS NOT NULL GROUP BY service.name union SELECT 'Total' AS 'Services' ,SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0 END) AS 'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 'Incident::Major', SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', SUM(case WHEN ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name ='Report'  THEN 1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0 END) AS 'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk',SUM(case WHEN ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN ( 'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem', 'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'Total' FROM ticket, service, ticket_type,ticket_state WHERE  ticket.service_id=service.id and ticket.ticket_state_id=ticket_state.id AND ticket.type_id=ticket_type.id and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged', 
                        'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
                        workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query29)
            result23 = cursor.fetchall()
            sql_query31 = """SELECT users.Login , SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  END) AS 
                        'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 'Incident::Major', 
                        SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', SUM(case WHEN 
                        ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name ='Report'  THEN 
                        1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  END) AS 
                        'Maintenance',SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
                        ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
                        ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
                        'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
                        'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'TOTAL' FROM ticket, users ,
                        ticket_type,ticket_state WHERE  ticket.user_id=users.id AND ticket.type_id=ticket_type.id AND ticket.sla_id 
                        IS NOT NULL and ticket.ticket_state_id=ticket_state.id and case when ticket_state.name IN ('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
                        reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when ticket_state.name IN ('CLOSED SUCCESSFUL',
                        'merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end 
                        GROUP BY users.Login union 
                        SELECT 'Total' AS 'Login',SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  END) AS 
                        'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 'Incident::Major', 
                        SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', SUM(case WHEN 
                        ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name ='Report'  THEN 
                        1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  END) AS 
                        'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
                        ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
                        ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
                        'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
                        'ServiceRequest','Incident::Major','Incident') THEN 1 ELSE 0 END) AS 'Total' FROM ticket, users ,
                        ticket_type,ticket_state WHERE  ticket.user_id=users.id AND ticket.type_id=ticket_type.id AND ticket.sla_id 
                        IS NOT NULL and ticket.ticket_state_id=ticket_state.id and case when ticket_state.name IN ('OPEN',
                        'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
                        reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when ticket_state.name IN ('CLOSED SUCCESSFUL',
                        'merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
                        'closed with workaround','RESOLVED') then  ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query31)
            result25 = cursor.fetchall()
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output)
            worksheet = workbook.add_worksheet()
            row = 0
            col = 0
            bold = workbook.add_format({'bold': True})
            worksheet.write('A1', 'Tickets', bold)
            worksheet.write('B1', 'Completed_Tickets', bold)

            row += 1

            for Tickets, Completed_Tickets in result15:
                worksheet.write(row, col, Tickets)
                worksheet.write(row, col + 1, Completed_Tickets)
                row += 1

            row1 = row + 3
            # worksheet.write('A' + str(row1), 'Summary Of Tickets_Type Wise',bold)
            worksheet.write('A' + str(row1), 'Name', bold)
            worksheet.write('B' + str(row1), 'Open', bold)
            worksheet.write('C' + str(row1), 'Closed', bold)
            worksheet.write('D' + str(row1), 'Total', bold)
            for Name, Open, Closed, Total in result17:
                worksheet.write(row1, col, Name)
                worksheet.write(row1, col + 1, Open)
                worksheet.write(row1, col + 2, Closed)
                worksheet.write(row1, col + 3, Total)
                row1 += 1

            row3 = row1 + 3
            worksheet.write('A' + str(row3), 'Queue', bold)
            worksheet.write('B' + str(row3), 'Open', bold)
            worksheet.write('C' + str(row3), 'Closed', bold)
            worksheet.write('D' + str(row3), 'Total', bold)

            for Queue, Open, Closed, Total in result19:
                worksheet.write(row3, col, Queue)
                worksheet.write(row3, col + 1, Open)
                worksheet.write(row3, col + 2, Closed)
                worksheet.write(row3, col + 3, Total)
                row3 += 1

            row5 = row3 + 3
            worksheet.write('A' + str(row5), 'Login', bold)
            worksheet.write('B' + str(row5), 'Open', bold)
            worksheet.write('C' + str(row5), 'Closed', bold)
            worksheet.write('D' + str(row5), 'Total', bold)
            for Login, Open, Closed, Total in result21:
                worksheet.write(row5, col, Login)
                worksheet.write(row5, col + 1, Open)
                worksheet.write(row5, col + 2, Closed)
                worksheet.write(row5, col + 3, Total)
                row5 += 1

            row7 = row5 + 3
            worksheet.write('A' + str(row7), 'Services', bold)
            worksheet.write('B' + str(row7), 'Incident', bold)
            worksheet.write('C' + str(row7), 'IncidentMajor', bold)
            worksheet.write('D' + str(row7), 'ServiceRequest', bold)
            worksheet.write('E' + str(row7), 'Problem', bold)
            worksheet.write('F' + str(row7), 'Report', bold)
            worksheet.write('G' + str(row7), 'Maintenance', bold)
            worksheet.write('H' + str(row7), 'Junk', bold)
            worksheet.write('I' + str(row7), 'Projects', bold)
            worksheet.write('J' + str(row7), 'Notification', bold)
            worksheet.write('K' + str(row7), 'Total', bold)
            for Services, Incident, IncidentMajor, ServiceRequest, Problem, Report, Maintenance, Junk, Projects, Notification, Total in result23:
                worksheet.write(row7, col, Services)
                worksheet.write(row7, col + 1, Incident)
                worksheet.write(row7, col + 2, IncidentMajor)
                worksheet.write(row7, col + 3, ServiceRequest)
                worksheet.write(row7, col + 4, Problem)
                worksheet.write(row7, col + 5, Report)
                worksheet.write(row7, col + 6, Maintenance)
                worksheet.write(row7, col + 7, Junk)
                worksheet.write(row7, col + 8, Projects)
                worksheet.write(row7, col + 9, Notification)
                worksheet.write(row7, col + 10, Total)
                row7 += 1

            row9 = row7 + 3
            worksheet.write('A' + str(row9), 'Login', bold)
            worksheet.write('B' + str(row9), 'Incident', bold)
            worksheet.write('C' + str(row9), 'IncidentMajor', bold)
            worksheet.write('D' + str(row9), 'ServiceRequest', bold)
            worksheet.write('E' + str(row9), 'Problem', bold)
            worksheet.write('F' + str(row9), 'Report', bold)
            worksheet.write('G' + str(row9), 'Maintenance', bold)
            worksheet.write('H' + str(row9), 'Junk', bold)
            worksheet.write('I' + str(row9), 'Projects', bold)
            worksheet.write('J' + str(row9), 'Notification', bold)
            worksheet.write('K' + str(row9), 'Total', bold)
            for Login, Incident, IncidentMajor, ServiceRequest, Problem, Report, Maintenance, Junk, Projects, Notification, Total in result25:
                worksheet.write(row9, col, Login)
                worksheet.write(row9, col + 1, Incident)
                worksheet.write(row9, col + 2, IncidentMajor)
                worksheet.write(row9, col + 3, ServiceRequest)
                worksheet.write(row9, col + 4, Problem)
                worksheet.write(row9, col + 5, Report)
                worksheet.write(row9, col + 6, Maintenance)
                worksheet.write(row9, col + 7, Junk)
                worksheet.write(row9, col + 8, Projects)
                worksheet.write(row9, col + 9, Notification)
                worksheet.write(row9, col + 10, Total)
                row9 += 1

            workbook.close()
            output.seek(0)
            filename = 'Full Summary Report.xlsx'
            response = HttpResponse(
                output,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=%s' % filename
            return response
        else:
            date1 = request.POST['date1']
            date2 = request.POST['date2']

            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query21 = """SELECT 'COMPLETED TICKETS' AS 'TICKETS', SUM(case WHEN ticket_state.name IN ('CLOSED 
            SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
            close-','closed with workaround','RESOLVED') THEN 1  END) AS 'COMPLETED TICKETS' FROM ticket_state,ticket, 
            queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS NOT NULL 
            and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor', 
            'ON-HOLD','Waiting for Customer','pending auto reopen') then  ticket.change_time <= '""" + date2 + """ 23:59:59' when 
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder', 
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end UNION SELECT 'OPENED TICKETS' AS 'TICKETS', SUM(case WHEN 
            ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD',
            'Waiting for Customer','pending auto reopen')  THEN 1 ELSE 0  END) AS 'OPEN TICKETS' FROM ticket_state,
            ticket,queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS 
            NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for 
            Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then ticket.create_time <= '""" + date2 + """ 23:59:59' 
            when ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end union 
            SELECT 'TOTAL NO. OF TICKETS' AS 'TICKETS',SUM(case WHEN ticket_state.name IN ('CLOSED 
            SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto 
            close-','closed with workaround','RESOLVED','pending auto reopen','Waiting for Customer','OPEN',
            'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD')THEN 1  END )AS 'TOTAL' FROM 
            ticket_state,ticket,queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND 
            ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then  
            ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end  """
            # print(sql_query21)
            cursor.execute(sql_query21)
            result15 = cursor.fetchall()
            # print(result15)
            sql_query23 = """SELECT ticket_type.Name,SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS',
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1  
            ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful',
            'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') 
            THEN 1  ELSE 0 END) AS 'Closed', SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
            'pending reminder', 'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED',
            'pending auto reopen', 'Waiting for Customer') THEN 1 ELSE 0 END) AS 'Total' FROM ticket,ticket_type,
            ticket_state WHERE ticket.type_id=ticket_type.id AND ticket.ticket_state_id=ticket_state.id AND ticket.sla_id 
            IS NOT NULL /*	AND ticket.change_time >= '""" + date1 + """' AND ticket.change_time <= '""" + date2 + """23
            :59:59' */ and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for 
            Vendor','ON-HOLD', 'Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when 
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end GROUP BY ticket_type.id union
            SELECT 'Total' as 'Name', SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS',
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1  
            ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful',
            'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') 
            THEN 1  ELSE 0 END) AS 'Closed', count(ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval',
            'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged','closed unsuccessful','removed',
            'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED',
            'pending auto reopen','Waiting for Customer')) as 'Total' FROM  ticket,ticket_type,ticket_state WHERE  	
            ticket.type_id=ticket_type.id AND  ticket.ticket_state_id=ticket_state.id AND ticket.sla_id IS NOT NULL /*	
            AND ticket.change_time >= '""" + date1 + """' AND ticket.change_time <= '""" + date2 + """ 23:59:59'*/ and case when 
            ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD',
            'Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when  
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query23)
            result17 = cursor.fetchall()
            sql_query25 = """SELECT  queue.Name AS Queue, SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS',
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
            ELSE 0  END) AS 'Open' , SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN('OPEN',
            'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED','pending auto reopen','Waiting for Customer')) AS 'Total' FROM ticket_state,ticket,
            queue WHERE ticket.ticket_state_id=ticket_state.id AND ticket.queue_id=queue.id AND ticket.sla_id IS NOT NULL 
            and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
            'ON-HOLD','Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when  
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder',
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end GROUP BY queue.Name union 
            SELECT 'Total' AS 'Queue', SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
            ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful', 'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN ('OPEN',
            'WORK IN PROGRESS','Waiting for Approval', 'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED', 'pending auto reopen','Waiting for Customer')) AS 'Total' FROM ticket,ticket_type,
            ticket_state WHERE ticket.type_id=ticket_type.id AND  ticket.ticket_state_id=ticket_state.id AND
            ticket.sla_id IS NOT NULL /* AND ticket.change_time >= '""" + date1 + """' AND ticket.change_time <= '""" + date2 + """ 23:59:59'*/ 
            and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor',
            'ON-HOLD', 'Waiting for Customer','pending auto reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:50' when 
            ticket_state.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder', 
            'PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') then  ticket.change_time 
            between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query25)
            result19 = cursor.fetchall()
            sql_query27 = """SELECT  users.Login , SUM(CASE WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
            ELSE 0  END) AS 'Open', SUM(CASE WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged', 
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') THEN 1 ELSE 0 END) AS 'Closed', SUM(CASE WHEN ticket_state.name IN ('OPEN',
            'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder', 'PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED','pending auto reopen', 'Waiting for Customer')  THEN 1 ELSE 0  END) AS 'Total' FROM 
            ticket, users ,ticket_state WHERE ticket.user_id=users.id AND ticket.ticket_state_id=ticket_state.id AND 
            ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor', 'ON-HOLD','Waiting for Customer','pending auto reopen') then  
            ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' 
            end GROUP BY users.Login union 
            SELECT 'Total' AS 'Login',SUM(case WHEN ticket_state.name IN ('OPEN','WORK IN PROGRESS', 
            'Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen')  THEN 1 
            ELSE 0 END) AS 'Open', SUM(case WHEN ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful', 'removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') THEN 1  ELSE 0 END) AS 'Closed', COUNT(ticket_state.name IN ('OPEN',
            'WORK IN PROGRESS','Waiting for Approval', 'Waiting for Vendor','ON-HOLD','CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed', 'pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED', 'pending auto reopen','Waiting for Customer')) AS 'Total' FROM    ticket,ticket_type,
            ticket_state WHERE ticket.type_id=ticket_type.id AND  ticket.ticket_state_id=ticket_state.id AND 
            ticket.sla_id IS NOT NULL and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then 
            ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL', 'merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE', 'pending auto close-','closed with 
            workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end"""
            cursor.execute(sql_query27)
            result21 = cursor.fetchall()
            sql_query29 = """SELECT  service.name AS Services , SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  
            END) AS 'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 
            'Incident::Major', SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', 
            SUM(case WHEN ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem',SUM(case WHEN ticket_type.name 
            ='Report'  THEN 1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  
            END) AS 'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
            ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
            ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
            'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
            'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'TOTAL' FROM ticket, service,
            ticket_type,ticket_state WHERE  ticket.service_id=service.id and ticket.ticket_state_id=ticket_state.id AND 
            ticket.type_id=ticket_type.id and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then  
            ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') then  ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end AND ticket.sla_id 
            IS NOT NULL GROUP BY service.name union
            SELECT 'Total' AS 'Services' ,SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  
            END) AS 'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 
            'Incident::Major', SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', 
            SUM(case WHEN ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name 
            ='Report'  THEN 1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  
            END) AS 'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk',SUM(case WHEN 
            ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
            ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
            'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
            'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'Total' FROM ticket, service,
            ticket_type,ticket_state WHERE  ticket.service_id=service.id and ticket.ticket_state_id=ticket_state.id AND 
            ticket.type_id=ticket_type.id and case when ticket_state.name IN ('OPEN','WORK IN PROGRESS','Waiting for 
            Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto reopen') then  
            ticket.create_time <= '""" + date2 + """ 23:59:59' when  ticket_state.name IN ('CLOSED SUCCESSFUL','merged',
            'closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with 
            workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query29)
            result23 = cursor.fetchall()
            sql_query31 = """SELECT users.Login , SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  END) AS 
            'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 'Incident::Major', 
            SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', SUM(case WHEN 
            ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name ='Report'  THEN 
            1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  END) AS 
            'Maintenance',SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
            ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
            ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
            'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
            'ServiceRequest','Incident::Major','Incident')  THEN 1 ELSE 0  END) AS 'TOTAL' FROM ticket, users ,
            ticket_type,ticket_state WHERE  ticket.user_id=users.id AND ticket.type_id=ticket_type.id AND ticket.sla_id 
            IS NOT NULL and ticket.ticket_state_id=ticket_state.id and case when ticket_state.name IN ('OPEN',
            'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
            reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when ticket_state.name IN ('CLOSED SUCCESSFUL',
            'merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') then ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end 
            GROUP BY users.Login union 
            SELECT 'Total' AS 'Login',SUM(case WHEN ticket_type.name ='Incident'  THEN 1 ELSE 0  END) AS 
            'Incident', SUM(case WHEN ticket_type.name ='Incident::Major'  THEN 1 ELSE 0  END) AS 'Incident::Major', 
            SUM(case WHEN ticket_type.name ='ServiceRequest'  THEN 1 ELSE 0  END) AS 'ServiceRequest', SUM(case WHEN 
            ticket_type.name ='Problem'  THEN 1 ELSE 0  END) AS 'Problem', SUM(case WHEN ticket_type.name ='Report'  THEN 
            1 ELSE 0  END) AS 'Report', SUM(case WHEN ticket_type.name ='Maintenance'  THEN 1 ELSE 0  END) AS 
            'Maintenance', SUM(case WHEN ticket_type.name ='Junk'  THEN 1 ELSE 0  END) AS 'Junk', SUM(case WHEN 
            ticket_type.name ='Projects & Oncall'  THEN 1 ELSE 0  END) AS 'Projects & Oncall', SUM(case WHEN 
            ticket_type.name ='Notification'  THEN 1 ELSE 0  END) AS 'Notification', SUM(case WHEN ticket_type.name IN (
            'Projects & Oncall','Notification','Junk','Follow-up','Maintenance','Report','RFC','Problem',
            'ServiceRequest','Incident::Major','Incident') THEN 1 ELSE 0 END) AS 'Total' FROM ticket, users ,
            ticket_type,ticket_state WHERE  ticket.user_id=users.id AND ticket.type_id=ticket_type.id AND ticket.sla_id 
            IS NOT NULL and ticket.ticket_state_id=ticket_state.id and case when ticket_state.name IN ('OPEN',
            'WORK IN PROGRESS','Waiting for Approval','Waiting for Vendor','ON-HOLD','Waiting for Customer','pending auto 
            reopen') then  ticket.create_time <= '""" + date2 + """ 23:59:59' when ticket_state.name IN ('CLOSED SUCCESSFUL',
            'merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-',
            'closed with workaround','RESOLVED') then  ticket.change_time between '""" + date1 + """' and '""" + date2 + """ 23:59:59' end """
            cursor.execute(sql_query31)
            result25 = cursor.fetchall()

            context10 = {

                'fsummary1': result15, 'fsummary3': result17,
                'fsummary5': result19, 'fsummary7': result21,
                'fsummary9': result23, 'fsummary11': result25,
                'hdate1': date1, 'hdate2': date2
            }
            return render(request, 'base2.html', context10)
    else:

        return render(request, 'fullsummaryreport.html')


def fullcustomer(request):
    if request.method == "POST":

        if 'Click Download' in request.POST:
            date1 = request.POST['date1']
            # print(date1)
            date2 = request.POST['date2']
            fullcustomer_name = request.POST['fullcustomer']
            print(fullcustomer_name)
            if fullcustomer_name:

                var2 = """ having Customer = '""" + fullcustomer_name + """' """
            else:
                var2 = ""
            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query32 = """SELECT  t.tn as Ticket_Id,t.title as Title,     (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
                        from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
                        and thi.create_time=(select max(thii.create_time) from ticket_history thii
                        where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
                        
                        DATE_FORMAT(t.create_time, '%Y-%m-%d') as create_date,DATE_FORMAT(t.create_time, '%H:%i:%s') as create_time,
								CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN date_format(t.change_time,'%Y-%m-%d') ELSE NULL END AS Closed_date,
								CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN date_format(t.change_time,'%H:%i:%s') ELSE NULL END AS Closed_time,
                        ts.name AS ticket_state,sl.name AS sla,tt.name Type_Name ,q.name AS queue_name,se.name AS service,
                        (select substring(cast(thi.name as char(100)),'29',position('%%OldValue%%' in thi.name)-29) as Customer
                        from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Source%'
                        and thi.create_time=(select max(thii.create_time) from ticket_history thii
                        where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') limit 1) as SOURCE,
                        (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
                        from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Customer%'
                        and thi.create_time=(select max(thii.create_time) from ticket_history thii
                        where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer
                        ,t.customer_id,t.customer_user_id,u.login as Responsible_user,
                        (select substring(cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent
                        from ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%'
                        and thi.create_time=(select max(thii.create_time) from ticket_history thii 
                        where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%') limit 1) as Time_Spent 
                        ,
                        CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age
                     

                        FROM 
                        ticket_state ts,
                        users u,
                        ticket t
                        LEFT JOIN ticket_type tt ON t.type_id = tt.id
                        LEFT JOIN sla sl ON t.sla_id = sl.id
                        LEFT JOIN service se ON t.service_id = se.id
                        LEFT JOIN queue q ON t.queue_id=q.id
                        WHERE
                        t.ticket_state_id =ts.id
                        AND t.user_id=u.id
                        AND tt.name NOT IN ('junk')
                        AND ts.name NOT IN ('merged')
                        AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
                        #AND u.login NOT IN ('root@localhost')  
                        AND t.create_time BETWEEN '""" + date1 + """' and '""" + date2 + """ 23:59:59' """+str(var2)+"""
                        #AND t.create_time BETWEEN '2023-01-01' AND '2023-01-25' 
                        ORDER BY t.create_time DESC;
                        """

            # print(sql_query32)
            cursor.execute(sql_query32)
            result26 = cursor.fetchall()
            # print(result26)
            

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output)
            worksheet = workbook.add_worksheet()
            row = 0
            col = 0
            # lst=['Responsible_user', 'Open', 'Closed', 'Total']
            # worksheet.set_column(lst)
            bold = workbook.add_format({'bold': True})
            lst = ['Ticket_Id','Title','Category','	create_date','	create_time','	Closed_date','	Closed_time	','State','	sla','	Type','Queue','	Service	','SOURCE','	Customer','	customer_id	','customer_user_id','	Responsible_user','	Time_Spent','Age']
            #lst = ['Ticket_Id', 'Responsible_user', 'Customer_id', 'Customer_user_id', 'Type_Name', 'Subject','Time_Spent', 'Customer', 'Created_Time', 'Closed_Time', 'Age',' Queue_Name', 'SLA', 'Service',' Ticket_State', 'Category']
            worksheet.write_row(0, 0, lst)
            row += 1

            # col20=0 var1 = Ticket_Id, Responsible_user, Customer_id, Customer_user_id, Type_Name, Subject,
            # Time_Spent, Customer, Created_Time, Closed_Time, Age, Queue_Name, SLA, Service, Ticket_State, Category
            for Ticket_Id, Responsible_user, Customer_id, Customer_user_id, Type_Name, Subject, Time_Spent, Customer, Created_Time, Closed_Time, Age, Queue_Name, SLA, Service, Ticket_State, Category, Source, Source1,Source2 in result26:
                worksheet.write(row, col, Ticket_Id)
                worksheet.write(row, col + 1, Responsible_user)
                worksheet.write(row, col + 2, Customer_id)
                worksheet.write(row, col + 3, Customer_user_id)
                worksheet.write(row, col + 4, Type_Name)
                worksheet.write(row, col + 5, Subject)
                worksheet.write(row, col + 6, Time_Spent)
                worksheet.write(row, col + 7, Customer)
                worksheet.write(row, col + 8, Created_Time)
                worksheet.write(row, col + 9, Closed_Time)
                worksheet.write(row, col + 10, Age)
                worksheet.write(row, col + 11, Queue_Name)
                worksheet.write(row, col + 12, SLA)
                worksheet.write(row, col + 13, Service)
                worksheet.write(row, col + 14, Ticket_State)
                worksheet.write(row, col + 15, Category)
                worksheet.write(row, col + 16, Source)
                worksheet.write(row, col + 17, Source1)
                worksheet.write(row, col + 18, Source2)
                row += 1
            workbook.close()
            output.seek(0)
            filename = 'Full Customer Report.xlsx'
            response = HttpResponse(
                output,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename=%s' % filename
            return response
        else:
            date1 = request.POST['date1']
            # print(date1)
            date2 = request.POST['date2']
            fullcustomer_name = request.POST['fullcustomer']
            
            if fullcustomer_name:

                var2 = """ having Customer = '""" + fullcustomer_name + """' """
            else:
                var2 = ""

            db_conn = mysql.connector.connect(host='otrs.futurenet.in', port=3306, user='readuser2',
                                              password='6FbUDa5VM',
                                              database='otrs5')
            cursor = db_conn.cursor()
            sql_query32 = """SELECT distinct t.tn as Ticket_Id,u.login as Responsible_user,t.customer_id,t.customer_user_id,
                                    tt.name Type_Name ,t.title as Subject,(select substring(cast(thi.name as char(100)),'32',position('%%OldValue%%' in thi.name)-32) as Time_Spent
                                    from ticket_history thi where thi.ticket_id=t.id and thi.name like '%TimeSpent%%%%'
                                    and thi.create_time=(select max(thii.create_time) from ticket_history thii 
                                    where thii.ticket_id=thi.ticket_id AND thii.name like '%TimeSpent%%%%') limit 1) as Time_Spent 
                                    ,(select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
                                    from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Customer%'
                                    and thi.create_time=(select max(thii.create_time) from ticket_history thii
                                    where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Customer%') limit 1) as Customer
                                    ,t.create_time as Created_Time,CASE WHEN ts.name IN ('CLOSED SUCCESSFUL','merged','closed unsuccessful','removed','pending reminder','PENDING AUTO CLOSE','pending auto close-','closed with workaround','RESOLVED') THEN t.change_time ELSE NULL END AS Closed_time,
                                    CASE WHEN ts.name IN ('WORK IN PROGRESS','OPEN','ON-HOLD','Waiting for Approval','Waiting for Vendor','Waiting for Customer','pending auto reopen') THEN DATEDIFF(NOW(),t.create_time) ELSE DATEDIFF(t.change_time,t.create_time) END as Age,
                                    q.name AS queue_name,sl.name AS sla,se.name AS service,ts.name AS ticket_state,
                                    (select substring(cast(thi.name as char(100)),'31',position('%%OldValue%%' in thi.name)-31) as Customer
                                    from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Category%'
                                    and thi.create_time=(select max(thii.create_time) from ticket_history thii
                                    where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Category%') limit 1) as Category,
                                    (select substring(cast(thi.name as char(100)),'29',position('%%OldValue%%' in thi.name)-29) as Customer
                                    from ticket_history thi where thi.ticket_id=t.id and thi.name like '%FieldName%Source%'
                                    and thi.create_time=(select max(thii.create_time) from ticket_history thii
                                    where thii.ticket_id=thi.ticket_id AND thii.name like '%FieldName%Source%') limit 1) as Source
                                    FROM 
                                    ticket_state ts,
                                    users u,
                                    ticket t
                                    LEFT JOIN ticket_type tt ON t.type_id = tt.id
                                    LEFT JOIN sla sl ON t.sla_id = sl.id
                                    LEFT JOIN service se ON t.service_id = se.id
                                    LEFT JOIN queue q ON t.queue_id=q.id
                                    WHERE
                                    t.ticket_state_id =ts.id
                                    AND t.user_id=u.id
                                    AND tt.name NOT IN ('junk')
                                    AND ts.name NOT IN ('merged')
                                    AND q.name NOT IN ('SALES','PRESALES','ODOOHELPDESK','ODOO','Postmaster','CUSTOMER-ALERTS','CUSTOMER-ALERTS::JASMIN','CUSTOMER-ALERTS::HLF','CUSTOMER-ALERTS::ABAN','CUSTOMER-ALERTS::FNET-MON','CUSTOMER-ALERTS::HINDU MISSION HOSPITAL','CUSTOMER-ALERTS::CRYSTALHR','CUSTOMER-ALERTS::BIZZ','CUSTOMER-ALERTS::TRANSFORMA')
                                    #AND u.login NOT IN ('root@localhost')  
                                    AND t.create_time BETWEEN '""" + date1 + """' and '""" + date2 + """ 23:59:59' """ +str(var2)+"""
                                    ORDER BY t.create_time DESC;"""
            # print(sql_query32)
            cursor.execute(sql_query32)
            result26 = cursor.fetchall()
            print(result26)

            context18 = {

                'fullcustomers': result26,
                'hdate1': date1, 'hdate2': date2, 'fullcustomer':fullcustomer_name
            }
            return render(request, 'base3.html', context18)
    else:

        context29 = {
            'members': results
        }
        return render(request, 'fullcustomer.html', context29)
