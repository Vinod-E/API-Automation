import mysql
from mysql import connector


class DB_Cleanup():
    # Client Question Randomization - 7590 - AT Tenant amsin server
    # Coding evaluation - 6082 - Automation Tenant amsin server
    # Server Question Randomization - 7528 - AT tenant amsin server
    # Static QP evaluation - 5282 - Automation Tenant amsin server
    # Random QP evaluation - 5365 - Automation Tenant amsin server
    # Timer verification - 7518 - AT tenant amsin server
    test_ids = [7590, 6082, 7528, 5282, 5365, 7518]
    # test_ids = [7590]
    host_ip = '35.154.36.218'
    db_name = 'appserver_core'
    login_name = 'qauser'
    pwd = 'qauser'
    try:
        conn = mysql.connector.connect(host=host_ip, database=db_name, user=login_name, password=pwd)
        mycursor = conn.cursor()
        for i in test_ids:
            i = str(i)
            mycursor.execute(
                'delete from test_result_infos where testresult_id in (select id from test_results where testuser_id in (select id from test_users where test_id = ' + i + ' and login_time is not null));')
            conn.commit()
            print("Test Result Info Deleted", i)
            mycursor.execute(
                'delete from test_results where testuser_id in (select id from test_users where test_id = ' + i + ' and login_time is not null);')
            conn.commit()
            print('Test result deleted ', i)
            mycursor.execute(
                'delete from candidate_scores where testuser_id in (select id from test_users where test_id = ' + i + ' and login_time is not null);')
            conn.commit()
            print('Candidate score deleted ', i)
            mycursor.execute(
                'delete from test_user_login_infos where testuser_id in (select id from test_users where test_id = ' + i + ' and login_time is not null);')
            conn.commit()
            print('Test user login info deleted ', i)
            mycursor.execute(
                "update test_users set login_time = NULL, log_out_time = NULL, status = 0, client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,config = NULL, client_system_info = NULL, total_score = NULL, percentage = NULL, eval_on = NULL, eval_by = NULL, eval_status = 'NotEvaluated', eval_task_id = NULL where test_id = '" + i + "';")
            conn.commit()
            print('Test user login time reset ', i)
        mycursor.close()
        print('Connection closed')
    except Exception as e:
        print(e)
    print('Executed')
