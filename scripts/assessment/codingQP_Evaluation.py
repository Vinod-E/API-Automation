from __future__ import absolute_import
import requests
import json
import unittest
import mysql
import xlrd
import xlwt
import time
from mysql import connector


class codingQP_Evaluation(unittest.TestCase):
    def test_codingQP_Evaluation(self):
        # try:
        #
        #     conn = mysql.connector.connect(host='35.154.36.218',
        #                                    database='appserver_core',
        #                                    user='qauser',
        #                                    password='qauser')
        #     mycursor = conn.cursor()
        #
        #     mycursor.execute('delete from test_results where testuser_id in (select id from test_users where test_id = 6082 and login_time is not null);')
        #     conn.commit()
        #     print('Test result deleted')
        #     mycursor.execute('delete from candidate_scores where testuser_id in (select id from test_users where test_id = 6082 and login_time is not null);')
        #     conn.commit()
        #     print('Candidate score deleted')
        #     mycursor.execute('delete from test_user_login_infos where testuser_id in (select id from test_users where test_id = 6082 and login_time is not null);')
        #     conn.commit()
        #     print('Test user login info deleted')
        #     mycursor.execute('update test_users set login_time = NULL, log_out_time = NULL, status = 0, client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,config = NULL, client_system_info = NULL, total_score = NULL, percentage = NULL, eval_by = NULL, eval_on = NULL, eval_status = NotEvaluated, eval_task_id = NULL where test_id = 6082;')
        #     conn.commit()
        #     print('Test user login time reseted')
        #     mycursor.close()
        #     print('Connection closed')
        # except Exception as e:
        #     print(e)
        # print('Executed')

        tenant_alias = "automation"

        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')

        # --------------------------------------------------------------------------------------------------------------
        # Read from Excel
        # --------------------------------------------------------------------------------------------------------------
        wb = xlrd.open_workbook("/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Input Data/Assessment/codingQP_Evaluation.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Evaluation_Check')
        sh1 = wb.sheet_by_index(0)
        rows = sh1.row_values(0)
        question_id_1 = int(rows[8])
        question_id_2 = int(rows[12])
        question_ids = [question_id_1, question_id_2]

        # --------------------------------------------------------------------------------------------------------------
        # Header printing in Output Excel
        # --------------------------------------------------------------------------------------------------------------
        ws.write(0, 0, "Candidate Id", self.__style0)
        ws.write(0, 1, "Test User Id", self.__style0)
        ws.write(0, 2, "Login Name", self.__style0)
        ws.write(0, 3, "Password", self.__style0)
        ws.write(0, 4, "Test Id", self.__style0)
        ws.write(0, 5, "Section 1 Id", self.__style0)
        ws.write(0, 6, "Q1 Lang", self.__style0)
        ws.write(0, 7, "Q1 Lang Id", self.__style0)
        ws.write(0, 8, question_id_1, self.__style0)
        ws.write(0, 9, "Section 2 Id", self.__style0)
        ws.write(0, 10, "Q2 Lang", self.__style0)
        ws.write(0, 11, "Q3 Lang Id", self.__style0)
        ws.write(0, 12, question_id_2, self.__style0)
        ws.write(0, 13, "expected_sec_1_total", self.__style0), ws.write(0, 14, "Actual_Sec_1_Total", self.__style0)
        ws.write(0, 15, "expected_sec_2_total", self.__style0), ws.write(0, 16, "Actual_Sec_2_Total", self.__style0)
        ws.write(0, 17, "expected_grp_1_total", self.__style0), ws.write(0, 18, "Actual_Grp_1_Total", self.__style0)
        ws.write(0, 19, "expected_grp_2_total", self.__style0), ws.write(0, 20, "Actual_Grp_2_Total", self.__style0)
        ws.write(0, 21, "expected_test_total", self.__style0), ws.write(0, 22, "Actual_test_Total", self.__style0)
        ws.write(0, 23, "Expected Percentage", self.__style0), ws.write(0, 24, "Actual Percentage", self.__style0)
        ws.write(0, 25, "Expected Q1 Marks", self.__style0), ws.write(0, 26, "Actual Q1 Marks", self.__style0)
        ws.write(0, 27, "Expected Q2 Marks", self.__style0), ws.write(0, 28, "Actual Q2 Marks", self.__style0)
        ws.write(0, 29, "Expected_Q1TC1_Marks", self.__style0), ws.write(0, 30, "Actual_Q1TC1_Marks", self.__style0)
        ws.write(0, 31, "Expected_Q1TC2_Marks", self.__style0), ws.write(0, 32, "Actual_Q1TC2_Marks", self.__style0)
        ws.write(0, 33, "Expected_Q1TC3_Marks", self.__style0), ws.write(0, 34, "Actual_Q1TC3_Marks", self.__style0)
        ws.write(0, 35, "Expected_Q1TC4_Marks", self.__style0), ws.write(0, 36, "Actual_Q1TC4_Marks", self.__style0)
        ws.write(0, 37, "Expected_Q1TC5_Marks", self.__style0), ws.write(0, 38, "Actual_Q1TC5_Marks", self.__style0)
        ws.write(0, 39, "Expected_Q1TC6_Marks", self.__style0), ws.write(0, 40, "Actual_Q1TC6_Marks", self.__style0)
        ws.write(0, 41, "Expected_Q1TC7_Marks", self.__style0), ws.write(0, 42, "Actual_Q1TC7_Marks", self.__style0)
        ws.write(0, 43, "Expected_Q1TC8_Marks", self.__style0), ws.write(0, 44, "Actual_Q1TC8_Marks", self.__style0)
        ws.write(0, 45, "Expected_Q1TC9_Marks", self.__style0), ws.write(0, 46, "Actual_Q1TC9_Marks", self.__style0)
        ws.write(0, 47, "Expected_Q1TC10_Marks", self.__style0), ws.write(0, 48, "Actual_Q1TC10_Marks", self.__style0)

        ws.write(0, 49, "Expected_Q2TC1_Marks", self.__style0), ws.write(0, 50, "Actual_Q2TC1_Marks", self.__style0)
        ws.write(0, 51, "Expected_Q2TC2_Marks", self.__style0), ws.write(0, 52, "Actual_Q2TC2_Marks", self.__style0)
        ws.write(0, 53, "Expected_Q2TC3_Marks", self.__style0), ws.write(0, 54, "Actual_Q2TC3_Marks", self.__style0)
        ws.write(0, 55, "Expected_Q2TC4_Marks", self.__style0), ws.write(0, 56, "Actual_Q2TC4_Marks", self.__style0)
        ws.write(0, 57, "Expected_Q2TC5_Marks", self.__style0), ws.write(0, 58, "Actual_Q2TC5_Marks", self.__style0)
        ws.write(0, 59, "Expected_Q2TC6_Marks", self.__style0), ws.write(0, 60, "Actual_Q2TC6_Marks", self.__style0)
        ws.write(0, 61, "Expected_Q2TC7_Marks", self.__style0), ws.write(0, 62, "Actual_Q2TC7_Marks", self.__style0)
        ws.write(0, 63, "Expected_Q2TC8_Marks", self.__style0), ws.write(0, 64, "Actual_Q2TC8_Marks", self.__style0)
        ws.write(0, 65, "Expected_Q2TC9_Marks", self.__style0), ws.write(0, 66, "Actual_Q2TC9_Marks", self.__style0)
        ws.write(0, 67, "Expected_Q2TC10_Marks", self.__style0), ws.write(0, 68, "Actual_Q2TC10_Marks", self.__style0)
        n = 1
        row_num = n
        while n < sh1.nrows:
            rows = sh1.row_values(row_num)
            candidate_id = rows[0]
            test_user_id = int(rows[1])
            login_name = rows[2]
            passwords = rows[3]
            test_id = int(rows[4])
            section_id_1 = int(rows[5])
            q1_lang = rows[6]
            q1_lang_id = rows[7]
            question_id_1_ans = rows[8]
            section_id_2 = int(rows[9])
            q2_lang = rows[10]
            q2_lang_id = rows[11]
            question_id_2_ans = rows[12]
            expected_sec_1_total = rows[13]
            expected_sec_2_total = rows[15]
            expected_grp_1_total = rows[17]
            expected_grp_2_total = rows[19]
            expected_test_total = rows[21]
            expected_percentage = rows[23]
            expected_q1_marks = rows[25]
            expected_q2_marks = rows[27]
            expected_q1_tc1_marks = rows[29]
            expected_q1_tc2_marks = rows[31]
            expected_q1_tc3_marks = rows[33]
            expected_q1_tc4_marks = rows[35]
            expected_q1_tc5_marks = rows[37]
            expected_q1_tc6_marks = rows[39]
            expected_q1_tc7_marks = rows[41]
            expected_q1_tc8_marks = rows[43]
            expected_q1_tc9_marks = rows[45]
            expected_q1_tc10_marks = rows[47]

            expected_q2_tc1_marks = rows[49]
            expected_q2_tc2_marks = rows[51]
            expected_q2_tc3_marks = rows[53]
            expected_q2_tc4_marks = rows[55]
            expected_q2_tc5_marks = rows[57]
            expected_q2_tc6_marks = rows[59]
            expected_q2_tc7_marks = rows[61]
            expected_q2_tc8_marks = rows[63]
            expected_q2_tc9_marks = rows[65]
            expected_q2_tc10_marks = rows[67]

            # ----------------------------------------------------------------------------------------------------------
            # Candidate data printing from Input data excel to Output data Excel
            # ----------------------------------------------------------------------------------------------------------
            ws.write(row_num, 0, candidate_id)
            ws.write(row_num, 1, test_user_id)
            ws.write(row_num, 2, login_name)
            ws.write(row_num, 3, passwords)
            ws.write(row_num, 4, test_id)
            ws.write(row_num, 5, section_id_1)
            ws.write(row_num, 6, q1_lang)
            ws.write(row_num, 7, q1_lang_id)
            ws.write(row_num, 8, question_id_1_ans)
            ws.write(row_num, 9, section_id_2)
            ws.write(row_num, 10, q2_lang)
            ws.write(row_num, 11, q2_lang_id)
            ws.write(row_num, 12, question_id_2_ans)
            ws.write(row_num, 13, expected_sec_1_total)
            ws.write(row_num, 15, expected_sec_2_total)
            ws.write(row_num, 17, expected_grp_1_total)
            ws.write(row_num, 19, expected_grp_2_total)
            ws.write(row_num, 21, expected_test_total)
            ws.write(row_num, 23, expected_percentage)
            ws.write(row_num, 25, expected_q1_marks)
            ws.write(row_num, 27, expected_q2_marks)
            ws.write(row_num, 29, expected_q1_tc1_marks)
            ws.write(row_num, 31, expected_q1_tc2_marks)
            ws.write(row_num, 33, expected_q1_tc3_marks)
            ws.write(row_num, 35, expected_q1_tc4_marks)
            ws.write(row_num, 37, expected_q1_tc5_marks)
            ws.write(row_num, 39, expected_q1_tc6_marks)
            ws.write(row_num, 41, expected_q1_tc7_marks)
            ws.write(row_num, 43, expected_q1_tc8_marks)
            ws.write(row_num, 45, expected_q1_tc9_marks)
            ws.write(row_num, 47, expected_q1_tc10_marks)
            ws.write(row_num, 49, expected_q2_tc1_marks)
            ws.write(row_num, 51, expected_q2_tc2_marks)
            ws.write(row_num, 53, expected_q2_tc3_marks)
            ws.write(row_num, 55, expected_q2_tc4_marks)
            ws.write(row_num, 57, expected_q2_tc5_marks)
            ws.write(row_num, 59, expected_q2_tc6_marks)
            ws.write(row_num, 61, expected_q2_tc7_marks)
            ws.write(row_num, 63, expected_q2_tc8_marks)
            ws.write(row_num, 65, expected_q2_tc9_marks)
            ws.write(row_num, 67, expected_q2_tc10_marks)

            # ----------------------------------------------------------------------------------------------------------
            # Login to HTML Test/Online Assessment
            # ----------------------------------------------------------------------------------------------------------
            login_to_test_header = {"content-type": "application/json"}
            login_to_test_data = {"ClientSystemInfo": "Browser:chrome/60.0.3112.78,OS:Linux x86_64,IPAddress:10.0.3.83",
                                  "IPAddress": "10.0.3.83", "IsOnlinePreview": False, "LoginName": login_name,
                                  "Password": passwords, "TenantAlias": tenant_alias}
            login_to_test_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/htmltest/api/v2/login_to_test/", headers=login_to_test_header,
                data=json.dumps(login_to_test_data), verify=True)
            self.login_to_test_response = login_to_test_request.json()
            self.login_to_test_token_val = self.login_to_test_response.get("Token")

            # ----------------------------------------------------------------------------------------------------------
            #  Submit Test API Call
            # ----------------------------------------------------------------------------------------------------------
            submit_test_result_header = {"content-type": "application/json",
                                         "X-AUTH-TOKEN": self.login_to_test_token_val}
            submit_test_result_data = {"isPartialSubmission": False, "disableBlockUI": False, "totalTimeSpent": 27,
                                       "testResultCollection": [
                                           {"q": question_id_1, "timeSpent": 8, "secId": section_id_1,
                                            "a": question_id_1_ans, "l": q1_lang_id},
                                           {"q": question_id_2, "timeSpent": 19, "secId": section_id_2,
                                            "a": question_id_2_ans, "l": q2_lang_id}],
                                       "config": "{\"TimeStamp\":\"2018-04-16T12:17:22.362Z\"}"}
            requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/submitTestResult/",
                          headers=submit_test_result_header,
                          data=json.dumps(submit_test_result_data, default=str), verify=True)

            # ----------------------------------------------------------------------------------------------------------
            #  Login to AMS
            # ----------------------------------------------------------------------------------------------------------
            crpo_login_header = {"content-type": "application/json"}
            data0 = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation",
                     "UserName": "admin"}
            response = requests.post('https://amsin.hirepro.in/py/common/user/login_user/', headers=crpo_login_header,
                                     data=json.dumps(data0), verify=True)
            self.TokenVal = response.json()
            self.NTokenVal = self.TokenVal.get("Token")

            # ----------------------------------------------------------------------------------------------------------
            #  Evaluate online assessment for candidate
            # ----------------------------------------------------------------------------------------------------------
            eval_online_assessment_header = {"content-type": "application/json",
                                             "X-AUTH-TOKEN": self.NTokenVal,
                                             "APP-NAME": "crpoassessment"}
            eval_online_assessment_data = {"testId": test_id, "candidateIds": [candidate_id]}
            # {"testId": 5926, "candidateIds": [1228284]}
            requests.post("https://amsin.hirepro.in/py/assessment/eval/api/v1/eval-online-assessment/",
                          headers=eval_online_assessment_header,
                          data=json.dumps(eval_online_assessment_data, default=str), verify=True)
            time.sleep(30)
            # ----------------------------------------------------------------------------------------------------------
            #  Fetch question wise candidate marks from DB and match with expected
            # ----------------------------------------------------------------------------------------------------------
            try:
                conn = mysql.connector.connect(host='35.154.36.218',
                                               database='appserver_core',
                                               user='hireprouser',
                                               password='tech@123')
                cursor = conn.cursor()
                cursor.execute("select id from Test_users where test_id = 6082 and candidate_id = %d;" % candidate_id)
                test_user_id = cursor.fetchone()
                cursor.execute("select question_id, obtained_marks from test_results where testuser_id = %d;" % test_user_id)
                data = cursor.fetchall()
                data_length = len(data)
                j = 0
                col = 26
                expected_cell = 25
                while j < len(question_ids):
                    i = 0
                    while i <= data_length:
                        if question_ids[j] == data[i][0]:
                            if rows[expected_cell] == data[i][1]:
                                ws.write(row_num, col, data[i][1], self.__style3)
                            else:
                                ws.write(row_num, col, data[i][1], self.__style2)
                                print("Question -- Candidate_Id - ", candidate_id, "Question_Id - ", question_ids[j],
                                      "Expected_Marks - ",
                                      rows[expected_cell], "Actual_Marks - ", data[i][1])
                            break
                        i += 1
                    j += 1
                    col += 2
                    expected_cell += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_1 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 29
                actual_tc_cell_pos = 30
                if len(test_results_question_marks) == 0:
                    for k in range(0, 10):
                        if round(rows[expected_tc_cell_pos], 2) == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {k[0]: k[1] for k in test_results_question_marks}
                    for k in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(k))
                        if round(rows[expected_tc_cell_pos], 2) == round(coding_question_vs_marks.get(k), 2):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(k), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(k), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_2 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 49
                actual_tc_cell_pos = 50
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 10):
                        if round(rows[expected_tc_cell_pos], 2) == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if round(rows[expected_tc_cell_pos], 2) == round(coding_question_vs_marks.get(kk), 2):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

            except Exception as e:
                print("Query error: Fetch question wise candidate marks from DB", e)
            finally:
                conn.close()
            # expected_tc_cell_pos = 29
            # actual_tc_cell_pos = 30
            # for k in coding_question_vs_marks:
            #     print(coding_question_vs_marks.get(k))
            #     if round(rows[expected_tc_cell_pos], 2) == round(coding_question_vs_marks.get(k), 2):
            #         ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(k), self.__style3)
            #     else:
            #         ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(k), self.__style2)
            #     expected_tc_cell_pos += 2
            #     actual_tc_cell_pos += 2

            # ----------------------------------------------------------------------------------------------------------
            #  View candidate scores by Candidate Id
            # ----------------------------------------------------------------------------------------------------------
            view_candidate_score_by_candidate_id_header = {"content-type": "application/json",
                                                           "X-AUTH-TOKEN": self.NTokenVal}
            view_candidate_score_by_candidate_id_data = {"TestId": str(int(test_id)),
                                                         "CandidateId": str(int(candidate_id)),
                                                         "TenantId": "ETg6fWphpuw="}  # "fhNePWEjLp8=" for at tenant
            print(view_candidate_score_by_candidate_id_data)
            view_candidate_score_by_candidate_id_request = requests.post(
                "https://amsin.hirepro.in/amsweb/JSONServices/JSONAssessmentManagementService.svc/ViewCandidateScoreByCandidateId",
                headers=view_candidate_score_by_candidate_id_header,
                data=json.dumps(view_candidate_score_by_candidate_id_data, default=str), verify=True)
            view_candidate_score_by_candidate_id_response = json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore']['TotalCandidateScore']
            print(view_candidate_score_by_candidate_id_response)

            # ----------------------------------------------------------------------------------------------------------
            #  Entering Actual data in excel and comparing Expected and Actual result
            # ----------------------------------------------------------------------------------------------------------
            actual_group1_score = 0
            actual_group2_score = 0
            k = 0
            while k <= 3:
                if json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13671:
                    actual_group1_score = \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k][
                            'Score']
                    if round(expected_grp_1_total, 3) == round(actual_group1_score, 3):
                        ws.write(row_num, 18, actual_group1_score, self.__style3)
                    else:
                        ws.write(row_num, 18, actual_group1_score, self.__style2)
                        group_1_id = json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidate_id, " group_1_id - ", group_1_id,
                              "Expected Score - ", expected_grp_1_total, "Actual Score - ", actual_group1_score)
                    break
                k += 1
            k = 0
            while k <= 3:
                if json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 48199:
                    actual_section1_score = \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k][
                            'Score']
                    if round(expected_sec_1_total, 3) == round(actual_section1_score, 3):
                        ws.write(row_num, 14, actual_section1_score, self.__style3)
                    else:
                        ws.write(row_num, 14, actual_section1_score, self.__style2)
                        section_1_id = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                                'TotalCandidateScore'][k][
                                'GroupId']
                        print("Section -- Candidate_Id - ", candidate_id, " section_1_id - ", section_1_id,
                              " Expected Score - ", expected_sec_1_total, " Actual Score - ", actual_section1_score)
                    break
                k += 1
            k = 0
            while k <= 3:
                if json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13672:
                    actual_group2_score = \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k][
                            'Score']
                    if round(expected_grp_2_total, 3) == round(actual_group2_score, 3):
                        ws.write(row_num, 20, actual_group2_score, self.__style3)
                    else:
                        ws.write(row_num, 20, actual_group2_score, self.__style2)
                        group_2_id = json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidate_id, " group_2_id - ", group_2_id,
                              "Expected Score - ", expected_grp_2_total, "Actual Score - ", actual_group2_score)
                    break
                k += 1
            k = 0
            while k <= 3:
                if json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 48200:
                    actual_section2_score = \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                            'TotalCandidateScore'][k][
                            'Score']
                    if round(expected_sec_2_total, 3) == round(actual_section2_score, 3):
                        ws.write(row_num, 16, actual_section2_score, self.__style3)
                    else:
                        ws.write(row_num, 16, actual_section2_score, self.__style2)
                        section_2_id = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['CandidateScore'][
                                'TotalCandidateScore'][k][
                                'GroupId']
                        print("Section -- Candidate_Id - ", candidate_id, " section_2_id - ", section_2_id,
                              " Expected Score - ", expected_sec_2_total, " Actual Score - ", actual_section2_score)
                    break
                k += 1
            actual_test_score = actual_group1_score + actual_group2_score
            if expected_test_total == actual_test_score:
                ws.write(row_num, 22, actual_test_score, self.__style3)
            else:
                ws.write(row_num, 22, actual_test_score, self.__style2)
                print("Section -- Candidate_Id - ", candidate_id, " Test Id - ", test_id, " Expected Test Score - ",
                      expected_test_total, " Actual Test Score - ", actual_test_score)
            wb_result.save("/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/codingQP_Evaluation_Check_New.xls")
            n += 1
            row_num += 1
