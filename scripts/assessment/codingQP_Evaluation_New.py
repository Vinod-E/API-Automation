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
        tenant_alias = "automation"
        sleep_Time = 30

        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.__style4 = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on')

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
        question_id_3 = int(rows[16])
        question_id_4 = int(rows[19])
        question_id_5 = int(rows[23])
        question_id_6 = int(rows[26])
        question_ids = [question_id_1, question_id_2, question_id_3, question_id_4, question_id_5, question_id_6]

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
        ws.write(0, 11, "Q2 Lang Id", self.__style0)
        ws.write(0, 12, question_id_2, self.__style0)
        ws.write(0, 13, "Section 3 Id", self.__style0)
        ws.write(0, 14, "Q3 Lang", self.__style0)
        ws.write(0, 15, "Q3 Lang Id", self.__style0)
        ws.write(0, 16, question_id_3, self.__style0)
        ws.write(0, 17, "Q4 Lang", self.__style0)
        ws.write(0, 18, "Q4 Lang Id", self.__style0)
        ws.write(0, 19, question_id_4, self.__style0)
        ws.write(0, 20, "Section 4 Id", self.__style0)
        ws.write(0, 21, "Q5 Lang", self.__style0)
        ws.write(0, 22, "Q5 Lang Id", self.__style0)
        ws.write(0, 23, question_id_5, self.__style0)
        ws.write(0, 24, "Q6 Lang", self.__style0)
        ws.write(0, 25, "Q6 Lang Id", self.__style0)
        ws.write(0, 26, question_id_6, self.__style0)

        ws.write(0, 27, "expected_sec_1_total", self.__style0), ws.write(0, 28, "Actual_Sec_1_Total", self.__style0)
        ws.write(0, 29, "expected_sec_2_total", self.__style0), ws.write(0, 30, "Actual_Sec_2_Total", self.__style0)
        ws.write(0, 31, "expected_sec_3_total", self.__style0), ws.write(0, 32, "Actual_Sec_3_Total", self.__style0)
        ws.write(0, 33, "expected_sec_4_total", self.__style0), ws.write(0, 34, "Actual_Sec_4_Total", self.__style0)

        ws.write(0, 35, "expected_grp_1_total", self.__style0), ws.write(0, 36, "Actual_Grp_1_Total", self.__style0)
        ws.write(0, 37, "expected_grp_2_total", self.__style0), ws.write(0, 38, "Actual_Grp_2_Total", self.__style0)

        ws.write(0, 39, "expected_test_total", self.__style0), ws.write(0, 40, "Actual_test_Total", self.__style0)

        ws.write(0, 41, "Expected Percentage", self.__style0), ws.write(0, 42, "X-GUID", self.__style0)

        ws.write(0, 43, "Expected Q1 Marks", self.__style0), ws.write(0, 44, "Actual Q1 Marks", self.__style0)
        ws.write(0, 45, "Expected Q2 Marks", self.__style0), ws.write(0, 46, "Actual Q2 Marks", self.__style0)
        ws.write(0, 47, "Expected Q3 Marks", self.__style0), ws.write(0, 48, "Actual Q3 Marks", self.__style0)
        ws.write(0, 49, "Expected Q4 Marks", self.__style0), ws.write(0, 50, "Actual Q4 Marks", self.__style0)
        ws.write(0, 51, "Expected Q5 Marks", self.__style0), ws.write(0, 52, "Actual Q5 Marks", self.__style0)
        ws.write(0, 53, "Expected Q6 Marks", self.__style0), ws.write(0, 54, "Actual Q6 Marks", self.__style0)

        ws.write(0, 55, "Expected_Q1TC1_Marks", self.__style0), ws.write(0, 56, "Actual_Q1TC1_Marks", self.__style0)

        ws.write(0, 57, "Expected_Q2TC1_Marks", self.__style0), ws.write(0, 58, "Actual_Q2TC1_Marks", self.__style0)
        ws.write(0, 59, "Expected_Q2TC2_Marks", self.__style0), ws.write(0, 60, "Actual_Q2TC2_Marks", self.__style0)
        ws.write(0, 61, "Expected_Q2TC3_Marks", self.__style0), ws.write(0, 62, "Actual_Q2TC3_Marks", self.__style0)
        ws.write(0, 63, "Expected_Q2TC4_Marks", self.__style0), ws.write(0, 64, "Actual_Q2TC4_Marks", self.__style0)
        ws.write(0, 65, "Expected_Q2TC5_Marks", self.__style0), ws.write(0, 66, "Actual_Q2TC5_Marks", self.__style0)
        ws.write(0, 67, "Expected_Q2TC6_Marks", self.__style0), ws.write(0, 68, "Actual_Q2TC6_Marks", self.__style0)
        ws.write(0, 69, "Expected_Q2TC7_Marks", self.__style0), ws.write(0, 70, "Actual_Q2TC7_Marks", self.__style0)
        ws.write(0, 71, "Expected_Q2TC8_Marks", self.__style0), ws.write(0, 72, "Actual_Q2TC8_Marks", self.__style0)
        ws.write(0, 73, "Expected_Q2TC9_Marks", self.__style0), ws.write(0, 74, "Actual_Q2TC9_Marks", self.__style0)
        ws.write(0, 75, "Expected_Q2TC10_Marks", self.__style0), ws.write(0, 76, "Actual_Q2TC10_Marks", self.__style0)

        ws.write(0, 77, "Expected_Q3TC1_Marks", self.__style0), ws.write(0, 78, "Actual_Q3TC1_Marks", self.__style0)

        ws.write(0, 79, "Expected_Q4TC1_Marks", self.__style0), ws.write(0, 80, "Actual_Q4TC1_Marks", self.__style0)

        ws.write(0, 81, "Expected_Q5TC1_Marks", self.__style0), ws.write(0, 82, "Actual_Q5TC1_Marks", self.__style0)
        ws.write(0, 83, "Expected_Q5TC2_Marks", self.__style0), ws.write(0, 84, "Actual_Q5TC2_Marks", self.__style0)
        ws.write(0, 85, "Expected_Q5TC3_Marks", self.__style0), ws.write(0, 86, "Actual_Q5TC3_Marks", self.__style0)
        ws.write(0, 87, "Expected_Q5TC4_Marks", self.__style0), ws.write(0, 88, "Actual_Q5TC4_Marks", self.__style0)
        ws.write(0, 89, "Expected_Q5TC5_Marks", self.__style0), ws.write(0, 90, "Actual_Q5TC5_Marks", self.__style0)
        ws.write(0, 91, "Expected_Q5TC6_Marks", self.__style0), ws.write(0, 92, "Actual_Q5TC6_Marks", self.__style0)
        ws.write(0, 93, "Expected_Q5TC7_Marks", self.__style0), ws.write(0, 94, "Actual_Q5TC7_Marks", self.__style0)
        ws.write(0, 95, "Expected_Q5TC8_Marks", self.__style0), ws.write(0, 96, "Actual_Q5TC8_Marks", self.__style0)
        ws.write(0, 97, "Expected_Q5TC9_Marks", self.__style0), ws.write(0, 98, "Actual_Q5TC9_Marks", self.__style0)
        ws.write(0, 99, "Expected_Q5TC10_Marks", self.__style0), ws.write(0, 100, "Actual_Q5TC10_Marks", self.__style0)

        ws.write(0, 101, "Expected_Q6TC1_Marks", self.__style0), ws.write(0, 102, "Actual_Q6TC1_Marks", self.__style0)

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

            section_id_3 = int(rows[13])
            q3_lang = rows[14]
            q3_lang_id = rows[15]
            question_id_3_ans = rows[16]
            q4_lang = rows[17]
            q4_lang_id = rows[18]
            question_id_4_ans = rows[19]

            section_id_4 = int(rows[20])
            q5_lang = rows[21]
            q5_lang_id = rows[22]
            question_id_5_ans = rows[23]
            q6_lang = rows[24]
            q6_lang_id = rows[25]
            question_id_6_ans = rows[26]

            expected_sec_1_total = rows[27]
            expected_sec_2_total = rows[29]
            expected_sec_3_total = rows[31]
            expected_sec_4_total = rows[33]

            expected_grp_1_total = rows[35]
            expected_grp_2_total = rows[37]

            expected_test_total = rows[39]

            expected_percentage = rows[41]

            expected_q1_marks = rows[43]
            expected_q2_marks = rows[45]
            expected_q3_marks = rows[47]
            expected_q4_marks = rows[49]
            expected_q5_marks = rows[51]
            expected_q6_marks = rows[53]

            expected_q1_tc1_marks = rows[55]

            expected_q2_tc1_marks = rows[57]
            expected_q2_tc2_marks = rows[59]
            expected_q2_tc3_marks = rows[61]
            expected_q2_tc4_marks = rows[63]
            expected_q2_tc5_marks = rows[65]
            expected_q2_tc6_marks = rows[67]
            expected_q2_tc7_marks = rows[69]
            expected_q2_tc8_marks = rows[71]
            expected_q2_tc9_marks = rows[73]
            expected_q2_tc10_marks = rows[75]

            expected_q3_tc1_marks = rows[77]

            expected_q4_tc1_marks = rows[79]

            expected_q5_tc1_marks = rows[81]
            expected_q5_tc2_marks = rows[83]
            expected_q5_tc3_marks = rows[85]
            expected_q5_tc4_marks = rows[87]
            expected_q5_tc5_marks = rows[89]
            expected_q5_tc6_marks = rows[91]
            expected_q5_tc7_marks = rows[93]
            expected_q5_tc8_marks = rows[95]
            expected_q5_tc9_marks = rows[97]
            expected_q5_tc10_marks = rows[99]

            expected_q6_tc1_marks = rows[101]

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
            ws.write(row_num, 13, section_id_3)
            ws.write(row_num, 14, q3_lang)
            ws.write(row_num, 15, q3_lang_id)
            ws.write(row_num, 16, question_id_3_ans)
            ws.write(row_num, 17, q4_lang)
            ws.write(row_num, 18, q4_lang_id)
            ws.write(row_num, 19, question_id_4_ans)
            ws.write(row_num, 20, section_id_4)
            ws.write(row_num, 21, q5_lang)
            ws.write(row_num, 22, q5_lang_id)
            ws.write(row_num, 23, question_id_5_ans)
            ws.write(row_num, 24, q6_lang)
            ws.write(row_num, 25, q6_lang_id)
            ws.write(row_num, 26, question_id_6_ans)

            ws.write(row_num, 27, expected_sec_1_total)
            ws.write(row_num, 29, expected_sec_2_total)
            ws.write(row_num, 31, expected_sec_3_total)
            ws.write(row_num, 33, expected_sec_4_total)

            ws.write(row_num, 35, expected_grp_1_total)
            ws.write(row_num, 37, expected_grp_2_total)

            ws.write(row_num, 39, expected_test_total)

            ws.write(row_num, 41, expected_percentage)

            ws.write(row_num, 43, expected_q1_marks)
            ws.write(row_num, 45, expected_q2_marks)
            ws.write(row_num, 47, expected_q3_marks)
            ws.write(row_num, 49, expected_q4_marks)
            ws.write(row_num, 51, expected_q5_marks)
            ws.write(row_num, 53, expected_q6_marks)

            ws.write(row_num, 55, expected_q1_tc1_marks)

            ws.write(row_num, 57, expected_q2_tc1_marks)
            ws.write(row_num, 59, expected_q2_tc2_marks)
            ws.write(row_num, 61, expected_q2_tc3_marks)
            ws.write(row_num, 63, expected_q2_tc4_marks)
            ws.write(row_num, 65, expected_q2_tc5_marks)
            ws.write(row_num, 67, expected_q2_tc6_marks)
            ws.write(row_num, 69, expected_q2_tc7_marks)
            ws.write(row_num, 71, expected_q2_tc8_marks)
            ws.write(row_num, 73, expected_q2_tc9_marks)
            ws.write(row_num, 75, expected_q2_tc10_marks)

            ws.write(row_num, 77, expected_q3_tc1_marks)

            ws.write(row_num, 79, expected_q4_tc1_marks)

            ws.write(row_num, 81, expected_q5_tc1_marks)
            ws.write(row_num, 83, expected_q5_tc2_marks)
            ws.write(row_num, 85, expected_q5_tc3_marks)
            ws.write(row_num, 87, expected_q5_tc4_marks)
            ws.write(row_num, 89, expected_q5_tc5_marks)
            ws.write(row_num, 91, expected_q5_tc6_marks)
            ws.write(row_num, 93, expected_q5_tc7_marks)
            ws.write(row_num, 95, expected_q5_tc8_marks)
            ws.write(row_num, 97, expected_q5_tc9_marks)
            ws.write(row_num, 99, expected_q5_tc10_marks)

            ws.write(row_num, 101, expected_q6_tc1_marks)

            # ----------------------------------------------------------------------------------------------------------
            # Login to HTML Test/Online Assessment
            # ----------------------------------------------------------------------------------------------------------
            login_to_test_header = {"content-type": "application/json", "APP-NAME": "onlineassessment", "X-APPLMA": "true"}
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
            submit_test_result_header = {"content-type": "application/json", "APP-NAME": "onlineassessment", "X-APPLMA": "true",
                                         "X-AUTH-TOKEN": self.login_to_test_token_val}
            submit_test_result_data = {"isPartialSubmission": False, "disableBlockUI": False, "totalTimeSpent": 27,
                                       "testResultCollection": [
                                           {"q": question_id_1, "timeSpent": 8, "secId": section_id_1,
                                            "a": question_id_1_ans, "l": q1_lang_id},
                                           {"q": question_id_2, "timeSpent": 19, "secId": section_id_2,
                                            "a": question_id_2_ans, "l": q2_lang_id},
                                           {"q": question_id_3, "timeSpent": 27, "secId": section_id_3,
                                            "a": question_id_3_ans, "l": q3_lang_id},
                                           {"q": question_id_4, "timeSpent": 48, "secId": section_id_3,
                                            "a": question_id_4_ans, "l": q4_lang_id},
                                           {"q": question_id_5, "timeSpent": 53, "secId": section_id_4,
                                            "a": question_id_5_ans, "l": q5_lang_id},
                                           {"q": question_id_6, "timeSpent": 58, "secId": section_id_4,
                                            "a": question_id_6_ans, "l": q6_lang_id}]}
            requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/submitTestResult/",
                          headers=submit_test_result_header,
                          data=json.dumps(submit_test_result_data, default=str), verify=True)

            # ----------------------------------------------------------------------------------------------------------
            #  Login to AMS
            # ----------------------------------------------------------------------------------------------------------
            crpo_login_header = {"content-type": "application/json", "APP-NAME": "crpo", "X-APPLMA": "true"}
            login_data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation",
                     "UserName": "admin"}
            response = requests.post('https://amsin.hirepro.in/py/common/user/login_user/', headers=crpo_login_header,
                                     data=json.dumps(login_data), verify=True)
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
            resp=requests.post("https://amsin.hirepro.in/py/assessment/eval/api/v1/eval-online-assessment/",
                          headers=eval_online_assessment_header,
                          data=json.dumps(eval_online_assessment_data, default=str), verify=True)
            GUID = resp.headers['X-GUID']
            ws.write(row_num, 42, GUID, self.__style4)
            time.sleep(sleep_Time)
            # ----------------------------------------------------------------------------------------------------------
            #  Fetch question wise candidate marks from DB and match with expected
            # ----------------------------------------------------------------------------------------------------------
            try:
                conn = mysql.connector.connect(host='35.154.36.218',
                                               database='appserver_core',
                                               user='qauser',
                                               password='qauser')
                cursor = conn.cursor()
                cursor.execute("select id from Test_users where test_id = 8581 and candidate_id = %d;" % candidate_id)
                test_user_id = cursor.fetchone()
                cursor.execute("select question_id, obtained_marks from test_results where testuser_id = %d;" % test_user_id)
                data = cursor.fetchall()
                data_length = len(data)
                j = 0
                col = 44
                expected_cell = 43
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
                expected_tc_cell_pos = 55
                actual_tc_cell_pos = 56
                if len(test_results_question_marks) == 0:
                    for k in range(0, 1):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, rows[expected_tc_cell_pos], self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_2 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 57
                actual_tc_cell_pos = 58
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 10):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, rows[expected_tc_cell_pos], self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_3 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 77
                actual_tc_cell_pos = 78
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 1):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, rows[expected_tc_cell_pos], self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_4 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 79
                actual_tc_cell_pos = 80
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 1):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, rows[expected_tc_cell_pos], self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_5 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 81
                actual_tc_cell_pos = 82
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 10):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, rows[expected_tc_cell_pos], self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2

                cursor.execute("select tri.coding_question_attachment_id testcase_id ,tri.coding_obtained_mark marks "
                               "from test_result_infos tri inner join test_results tr on tr.id = tri.testresult_id inner join test_users tu on tu.id = tr.testuser_id "
                               "where tu.test_id=%d" % test_id + " and tr.testuser_id = %d" % test_user_id + " and tr.question_id in (%s)" % question_id_6 + ";")
                test_results_question_marks = cursor.fetchall()
                print(test_results_question_marks)
                expected_tc_cell_pos = 101
                actual_tc_cell_pos = 102
                if len(test_results_question_marks) == 0:
                    for kk in range(0, 1):
                        if rows[expected_tc_cell_pos] == 0:
                            ws.write(row_num, actual_tc_cell_pos, 0, self.__style3)
                        elif rows[expected_tc_cell_pos] == 'Empty':
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, 'Empty', self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
                else:
                    coding_question_vs_marks = {kk[0]: kk[1] for kk in test_results_question_marks}
                    for kk in coding_question_vs_marks:
                        print(coding_question_vs_marks.get(kk))
                        if rows[expected_tc_cell_pos] == coding_question_vs_marks.get(kk):
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style3)
                        else:
                            ws.write(row_num, actual_tc_cell_pos, coding_question_vs_marks.get(kk), self.__style2)
                        expected_tc_cell_pos += 2
                        actual_tc_cell_pos += 2
            except Exception as e:
                print("Query error: Fetch question wise candidate marks from DB", e)
            finally:
                conn.close()

            # ----------------------------------------------------------------------------------------------------------
            #  View candidate scores by Candidate Id
            # ----------------------------------------------------------------------------------------------------------
            crpo_login_header = {"content-type": "application/json", "APP-NAME": "crpo", "X-APPLMA": "true"}
            login_data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation",
                     "UserName": "admin"}
            response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=crpo_login_header,
                                     data=json.dumps(login_data), verify=True)
            self.TokenVal = response.json()
            self.NTokenVal = self.TokenVal.get("Token")

            view_candidate_score_by_candidate_id_header = {"content-type": "application/json",
                                                           "X-AUTH-TOKEN": self.NTokenVal}
            view_candidate_score_by_candidate_id_data = {"testId": test_id, "candidateId": candidate_id,
                                                         "reportFlags": {'testUsersScoreRequired': True,
                                                                         'fileContentRequired': False}, "print": False}

            print(view_candidate_score_by_candidate_id_data)
            view_candidate_score_by_candidate_id_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/report/api/v1/candidatetranscript/",
                headers=view_candidate_score_by_candidate_id_header,
                data=json.dumps(view_candidate_score_by_candidate_id_data, default=str), verify=True)
            print(view_candidate_score_by_candidate_id_request)
            actual_group1_score = 0
            actual_group2_score = 0
            testUserScore = len(json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'])
            for ite in range (0, testUserScore):
                if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['candidateId'] == candidate_id:
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['groupId'] == 17180:
                            actual_group1_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['score']
                            if round(expected_grp_1_total, 3) == round(actual_group1_score, 3):
                                ws.write(row_num, 36, actual_group1_score, self.__style3)
                            else:
                                ws.write(row_num, 36, actual_group1_score, self.__style2)
                                group_1_id = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['groupId']
                                print("Group -- Candidate_Id - ", candidate_id, " group_1_id - ", group_1_id, "Expected Score - ", expected_grp_1_total, "Actual Score - ", actual_group1_score)
                            break
                        k += 1
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"] == 53421:
                            actual_section1_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]['score']
                            if round(expected_sec_1_total, 3) == round(actual_section1_score, 3):
                                ws.write(row_num, 28, actual_section1_score, self.__style3)
                            else:
                                ws.write(row_num, 28, actual_section1_score, self.__style2)
                                section_1_id = \
                                    json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidate_id, " section_1_id - ", section_1_id,
                                      "Expected Score - ", expected_sec_1_total, " Actual Score - ", actual_section1_score)
                            break
                        k += 1
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"] == 53422:
                            actual_section2_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]['score']
                            if round(expected_sec_2_total, 3) == round(actual_section2_score, 3):
                                ws.write(row_num, 30, actual_section2_score, self.__style3)
                            else:
                                ws.write(row_num, 30, actual_section2_score, self.__style2)
                                section_2_id = \
                                    json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidate_id, " section_2_id - ", section_2_id,
                                      "Expected Score - ", expected_sec_2_total, " Actual Score - ", actual_section2_score)
                            break
                        k += 1
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['groupId'] == 17181:
                            actual_group2_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['score']
                            if round(expected_grp_2_total, 3) == round(actual_group2_score, 3):
                                ws.write(row_num, 38, actual_group2_score, self.__style3)
                            else:
                                ws.write(row_num, 38, actual_group2_score, self.__style2)
                                group_2_id = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['groupId']
                                print("Group -- Candidate_Id - ", candidate_id, " group_2_id - ", group_2_id,
                                      "Expected Score - ", expected_grp_2_total, "Actual Score - ", actual_group2_score)
                            break
                        k += 1
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"] == 53423:
                            actual_section3_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]['score']
                            if round(expected_sec_3_total, 3) == round(actual_section3_score, 3):
                                ws.write(row_num, 32, actual_section3_score, self.__style3)
                            else:
                                ws.write(row_num, 32, actual_section3_score, self.__style2)
                                section_3_id = \
                                    json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidate_id, " section_3_id - ", section_3_id,
                                      "Expected Score - ", expected_sec_3_total, " Actual Score - ", actual_section3_score)
                            break
                        k += 1
                    k = 0
                    while k <= 6:
                        if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"] == 53424:
                            actual_section4_score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]['score']
                            if round(expected_sec_4_total, 3) == round(actual_section4_score, 3):
                                ws.write(row_num, 34, actual_section4_score, self.__style3)
                            else:
                                ws.write(row_num, 34, actual_section4_score, self.__style2)
                                section_4_id = \
                                    json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidate_id, " section_4_id - ", section_4_id,
                                      "Expected Score - ", expected_sec_4_total, " Actual Score - ", actual_section4_score)
                            break
                        k += 1

            # ----------------------------------------------------------------------------------------------------------
            #  Entering Actual data in excel and comparing Expected and Actual result
            # ----------------------------------------------------------------------------------------------------------

            actual_test_score = actual_group1_score + actual_group2_score
            if expected_test_total == actual_test_score:
                ws.write(row_num, 40, actual_test_score, self.__style3)
            else:
                ws.write(row_num, 40, actual_test_score, self.__style2)
                print("Section -- Candidate_Id - ", candidate_id, " Test Id - ", test_id, " Expected Test Score - ",
                      expected_test_total, " Actual Test Score - ", actual_test_score)
            wb_result.save("/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/CodingQP_Evaluation_Check.xls")
            n += 1
            row_num += 1
