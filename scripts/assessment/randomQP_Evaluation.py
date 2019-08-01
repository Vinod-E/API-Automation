from __future__ import absolute_import
import requests
import json
import unittest
import mysql
import xlrd
import xlwt
from mysql import connector
from itertools import combinations


class randomQP_Evaluation(unittest.TestCase):
    def test_randomQP_Evaluation(self):

        # try:
        #
        #     conn = mysql.connector.connect(host='35.154.36.218',
        #                                    database='appserver_core',
        #                                    user='qauser',
        #                                    password='qauser')
        #     mycursor = conn.cursor()
        #
        #     mycursor.execute(
        #         'delete from test_results where testuser_id in (select id from test_users where test_id = 5365 and login_time is not null);')
        #     conn.commit()
        #     print('Test result deleted')
        #     mycursor.execute(
        #         'delete from candidate_scores where testuser_id in (select id from test_users where test_id = 5365 and login_time is not null);')
        #     conn.commit()
        #     print('Candidate score deleted')
        #     mycursor.execute(
        #         'delete from test_user_login_infos where testuser_id in (select id from test_users where test_id = 5365 and login_time is not null);')
        #     conn.commit()
        #     print('Test user login info deleted')
        #     mycursor.execute(
        #         'update test_users set login_time = NULL, log_out_time = NULL, status = 0, client_system_info = NULL, time_spent = NULL, is_password_disabled = 0,config = NULL, client_system_info = NULL, total_score = NULL, percentage = NULL where test_id = 5365 and login_time is not null;')
        #     conn.commit()
        #     print('Test user login time reseted')
        #     mycursor.close()
        #     print('Connection closed')
        # except Exception as e:
        #     print(e)
        # print('Executed')
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
        wb = xlrd.open_workbook("/home/testingteam/hirepro_automation/API-Automation/Input Data/Assessment/staticrandomQP_Evaluation.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Evaluation_Check', cell_overwrite_ok=True)
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        sh1 = wb.sheet_by_index(0)
        rows = sh1.row_values(0)
        QuestionId_1 = int(rows[5])
        QuestionId_2 = int(rows[6])
        QuestionId_3 = int(rows[8])
        QuestionId_4 = int(rows[9])
        QuestionId_5 = int(rows[11])
        QuestionId_6 = int(rows[12])
        QuestionId_7 = int(rows[13])
        QuestionId_8 = int(rows[14])
        QuestionId_9 = int(rows[15])
        QuestionId_10 = int(rows[16])
        QuestionId_11 = int(rows[17])
        QuestionId_12 = int(rows[18])
        QuestionId_13 = int(rows[19])
        QuestionId_14 = int(rows[20])
        QuestionId_15 = int(rows[21])
        QuestionId_16 = int(rows[22])
        QuestionId_17 = int(rows[24])
        QuestionId_18 = int(rows[25])
        QuestionId_19 = int(rows[26])
        QuestionId_20 = int(rows[27])
        QuestionId_21 = int(rows[29])
        QuestionId_22 = int(rows[30])
        QuestionId_23 = int(rows[31])
        QuestionId_24 = int(rows[32])
        QuestionId_25 = int(rows[34])
        QuestionId_26 = int(rows[35])
        QuestionId_27 = int(rows[36])
        QuestionId_28 = int(rows[37])
        QuestionId_29 = int(rows[38])
        QuestionId_30 = int(rows[39])
        QuestionId_31 = int(rows[41])
        QuestionId_32 = int(rows[42])
        QuestionId_33 = int(rows[43])
        QuestionId_34 = int(rows[45])
        QuestionId_35 = int(rows[46])
        QuestionId_36 = int(rows[47])
        QuestionId_37 = int(rows[49])
        QuestionId_38 = int(rows[50])
        QuestionId_39 = int(rows[51])
        question_id_headers = [QuestionId_1, QuestionId_2, QuestionId_3, QuestionId_4, QuestionId_5, QuestionId_6,
                               QuestionId_7,
                               QuestionId_8, QuestionId_9, QuestionId_10, QuestionId_11, QuestionId_12, QuestionId_13,
                               QuestionId_14, QuestionId_15, QuestionId_16, QuestionId_17, QuestionId_18, QuestionId_19,
                               QuestionId_20, QuestionId_21, QuestionId_22, QuestionId_23, QuestionId_24, QuestionId_25,
                               QuestionId_26, QuestionId_27, QuestionId_28, QuestionId_29, QuestionId_30, QuestionId_31,
                               QuestionId_32, QuestionId_33, QuestionId_34, QuestionId_35, QuestionId_36, QuestionId_37,
                               QuestionId_38, QuestionId_39]

        Expected_Q1_Mark = int(rows[82])
        Expected_Q2_Mark = int(rows[84])
        Expected_Q3_Mark = int(rows[86])
        Expected_Q4_Mark = int(rows[88])
        Expected_Q5_Mark = int(rows[90])
        Expected_Q6_Mark = int(rows[92])
        Expected_Q7_Mark = int(rows[94])
        Expected_Q8_Mark = int(rows[96])
        Expected_Q9_Mark = int(rows[98])
        Expected_Q10_Mark = int(rows[100])
        Expected_Q11_Mark = int(rows[102])
        Expected_Q12_Mark = int(rows[104])
        Expected_Q13_Mark = int(rows[106])
        Expected_Q14_Mark = int(rows[108])
        Expected_Q15_Mark = int(rows[110])
        Expected_Q16_Mark = int(rows[112])
        Expected_Q17_Mark = int(rows[114])
        Expected_Q18_Mark = int(rows[116])
        Expected_Q19_Mark = int(rows[118])
        Expected_Q20_Mark = int(rows[120])
        Expected_Q21_Mark = int(rows[122])
        Expected_Q22_Mark = int(rows[124])
        Expected_Q23_Mark = int(rows[126])
        Expected_Q24_Mark = int(rows[128])
        Expected_Q25_Mark = int(rows[130])
        Expected_Q26_Mark = int(rows[132])
        Expected_Q27_Mark = int(rows[134])
        Expected_Q28_Mark = int(rows[136])
        Expected_Q29_Mark = int(rows[138])
        Expected_Q30_Mark = int(rows[140])
        Expected_Q31_Mark = int(rows[142])
        Expected_Q32_Mark = int(rows[144])
        Expected_Q33_Mark = int(rows[146])
        Expected_Q34_Mark = int(rows[148])
        Expected_Q35_Mark = int(rows[150])
        Expected_Q36_Mark = int(rows[152])
        Expected_Q37_Mark = int(rows[154])
        Expected_Q38_Mark = int(rows[156])
        Expected_Q39_Mark = int(rows[158])

        Expected_Sec_1_Total_Mark = int(rows[52])
        Expected_Sec_2_Total_Mark = int(rows[54])
        Expected_Sec_3_Total_Mark = int(rows[56])
        Expected_Sec_4_Total_Mark = int(rows[58])
        Expected_Sec_5_Total_Mark = int(rows[60])
        Expected_Sec_6_Total_Mark = int(rows[62])
        Expected_Sec_7_Total_Mark = int(rows[64])
        Expected_Sec_8_Total_Mark = int(rows[66])
        Expected_Sec_9_Total_Mark = int(rows[68])

        question_ids = [QuestionId_1, QuestionId_2, QuestionId_3, QuestionId_4, QuestionId_5, QuestionId_6, QuestionId_7,
                               QuestionId_8, QuestionId_9, QuestionId_10, QuestionId_11, QuestionId_12, QuestionId_13,
                               QuestionId_14, QuestionId_15, QuestionId_16, QuestionId_17, QuestionId_18, QuestionId_19,
                               QuestionId_20, QuestionId_21, QuestionId_22, QuestionId_23, QuestionId_24, QuestionId_25,
                               QuestionId_26, QuestionId_27, QuestionId_28, QuestionId_29, QuestionId_30, QuestionId_31,
                               QuestionId_32, QuestionId_33, QuestionId_34, QuestionId_35, QuestionId_36, QuestionId_37,
                               QuestionId_38, QuestionId_39]

        # --------------------------------------------------------------------------------------------------------------
        # Header printing in Output Excel
        # --------------------------------------------------------------------------------------------------------------
        ws.write(0, 0, "Candidate Id", self.__style0)
        ws.write(0, 1, "Login Name", self.__style0)
        ws.write(0, 2, "Password", self.__style0)
        ws.write(0, 3, "Test Id", self.__style0)
        ws.write(0, 4, "Section 1 Id", self.__style0)
        ws.write(0, 5, QuestionId_1, self.__style0)
        ws.write(0, 6, QuestionId_2, self.__style0)
        ws.write(0, 7, "Section 2 Id", self.__style0)
        ws.write(0, 8, QuestionId_3, self.__style0)
        ws.write(0, 9, QuestionId_4, self.__style0)
        ws.write(0, 10, "Section 3 Id", self.__style0)
        ws.write(0, 11, QuestionId_5, self.__style0)
        ws.write(0, 12, QuestionId_6, self.__style0)
        ws.write(0, 13, QuestionId_7, self.__style0)
        ws.write(0, 14, QuestionId_8, self.__style0)
        ws.write(0, 15, QuestionId_9, self.__style0)
        ws.write(0, 16, QuestionId_10, self.__style0)
        ws.write(0, 17, QuestionId_11, self.__style0)
        ws.write(0, 18, QuestionId_12, self.__style0)
        ws.write(0, 19, QuestionId_13, self.__style0)
        ws.write(0, 20, QuestionId_14, self.__style0)
        ws.write(0, 21, QuestionId_15, self.__style0)
        ws.write(0, 22, QuestionId_16, self.__style0)
        ws.write(0, 23, "Section 4 Id", self.__style0)
        ws.write(0, 24, QuestionId_17, self.__style0)
        ws.write(0, 25, QuestionId_18, self.__style0)
        ws.write(0, 26, QuestionId_19, self.__style0)
        ws.write(0, 27, QuestionId_20, self.__style0)
        ws.write(0, 28, "Section 5 Id", self.__style0)
        ws.write(0, 29, QuestionId_21, self.__style0)
        ws.write(0, 30, QuestionId_22, self.__style0)
        ws.write(0, 31, QuestionId_23, self.__style0)
        ws.write(0, 32, QuestionId_24, self.__style0)
        ws.write(0, 33, "Section 6 Id", self.__style0)
        ws.write(0, 34, QuestionId_25, self.__style0)
        ws.write(0, 35, QuestionId_26, self.__style0)
        ws.write(0, 36, QuestionId_27, self.__style0)
        ws.write(0, 37, QuestionId_28, self.__style0)
        ws.write(0, 38, QuestionId_29, self.__style0)
        ws.write(0, 39, QuestionId_30, self.__style0)
        ws.write(0, 40, "Section 7 Id", self.__style0)
        ws.write(0, 41, QuestionId_31, self.__style0)
        ws.write(0, 42, QuestionId_32, self.__style0)
        ws.write(0, 43, QuestionId_33, self.__style0)
        ws.write(0, 44, "Section 8 Id", self.__style0)
        ws.write(0, 45, QuestionId_34, self.__style0)
        ws.write(0, 46, QuestionId_35, self.__style0)
        ws.write(0, 47, QuestionId_36, self.__style0)
        ws.write(0, 48, "Section 9 Id", self.__style0)
        ws.write(0, 49, QuestionId_37, self.__style0)
        ws.write(0, 50, QuestionId_38, self.__style0)
        ws.write(0, 51, QuestionId_39, self.__style0)

        ws.write(0, 52, "Expected_Sec_1_Total", self.__style0)
        ws.write(0, 53, "Actual_Sec_1_Total", self.__style0)
        ws.write(0, 54, "Expected_Sec_2_Total", self.__style0)
        ws.write(0, 55, "Actual_Sec_2_Total", self.__style0)
        ws.write(0, 56, "Expected_Sec_3_Total", self.__style0)
        ws.write(0, 57, "Actual_Sec_3_Total", self.__style0)
        ws.write(0, 58, "Expected_Sec_4_Total", self.__style0)
        ws.write(0, 59, "Actual_Sec_4_Total", self.__style0)
        ws.write(0, 60, "Expected_Sec_5_Total", self.__style0)
        ws.write(0, 61, "Actual_Sec_5_Total", self.__style0)
        ws.write(0, 62, "Expected_Sec_6_Total", self.__style0)
        ws.write(0, 63, "Actual_Sec_6_Total", self.__style0)
        ws.write(0, 64, "Expected_Sec_7_Total", self.__style0)
        ws.write(0, 65, "Actual_Sec_7_Total", self.__style0)
        ws.write(0, 66, "Expected_Sec_8_Total", self.__style0)
        ws.write(0, 67, "Actual_Sec_8_Total", self.__style0)
        ws.write(0, 68, "Expected_Sec_9_Total", self.__style0)
        ws.write(0, 69, "Actual_Sec_9_Total", self.__style0)



        ws.write(0, 70, "Expected_Grp_1_Total", self.__style0)
        ws.write(0, 71, "Actual_Grp_1_Total", self.__style0)
        ws.write(0, 72, "Expected_Grp_2_Total", self.__style0)
        ws.write(0, 73, "Actual_Grp_2_Total", self.__style0)
        ws.write(0, 74, "Expected_Grp_3_Total", self.__style0)
        ws.write(0, 75, "Actual_Grp_3_Total", self.__style0)
        ws.write(0, 76, "Expected_Grp_4_Total", self.__style0)
        ws.write(0, 77, "Actual_Grp_4_Total", self.__style0)



        ws.write(0, 78, "Expected_test_Total", self.__style0)
        ws.write(0, 79, "Actual_test_Total", self.__style0)
        ws.write(0, 80, "Expected Percentage", self.__style0)
        ws.write(0, 81, "Actual Percentage", self.__style0)
        ws.write(0, 82, "Expected Q1 Marks", self.__style0)
        ws.write(0, 83, "Actual Q1 Marks", self.__style0)
        ws.write(0, 84, "Expected Q2 Marks", self.__style0)
        ws.write(0, 85, "Actual Q2 Marks", self.__style0)
        ws.write(0, 86, "Expected Q3 Marks", self.__style0)
        ws.write(0, 87, "Actual Q3 Marks", self.__style0)
        ws.write(0, 88, "Expected Q4 Marks", self.__style0)
        ws.write(0, 89, "Actual Q4 Marks", self.__style0)
        ws.write(0, 90, "Expected Q5 Marks", self.__style0)
        ws.write(0, 91, "Actual Q5 Marks", self.__style0)
        ws.write(0, 92, "Expected Q6 Marks", self.__style0)
        ws.write(0, 93, "Actual Q6 Marks", self.__style0)
        ws.write(0, 94, "Expected Q7 Marks", self.__style0)
        ws.write(0, 95, "Actual Q7 Marks", self.__style0)
        ws.write(0, 96, "Expected Q8 Marks", self.__style0)
        ws.write(0, 97, "Actual Q8 Marks", self.__style0)
        ws.write(0, 98, "Expected Q9 Marks", self.__style0)
        ws.write(0, 99, "Actual Q9 Marks", self.__style0)
        ws.write(0, 100, "Expected Q10 Marks", self.__style0)
        ws.write(0, 101, "Actual Q10 Marks", self.__style0)
        ws.write(0, 102, "Expected Q11 Marks", self.__style0)
        ws.write(0, 103, "Actual Q11 Marks", self.__style0)
        ws.write(0, 104, "Expected Q12 Marks", self.__style0)
        ws.write(0, 105, "Actual Q12 Marks", self.__style0)
        ws.write(0, 106, "Expected Q13 Marks", self.__style0)
        ws.write(0, 107, "Actual Q13 Marks", self.__style0)
        ws.write(0, 108, "Expected Q14 Marks", self.__style0)
        ws.write(0, 109, "Actual Q14 Marks", self.__style0)
        ws.write(0, 110, "Expected Q15 Marks", self.__style0)
        ws.write(0, 111, "Actual Q15 Marks", self.__style0)
        ws.write(0, 112, "Expected Q16 Marks", self.__style0)
        ws.write(0, 113, "Actual Q16 Marks", self.__style0)
        ws.write(0, 114, "Expected Q17 Marks", self.__style0)
        ws.write(0, 115, "Actual Q17 Marks", self.__style0)
        ws.write(0, 116, "Expected Q18 Marks", self.__style0)
        ws.write(0, 117, "Actual Q18 Marks", self.__style0)
        ws.write(0, 118, "Expected Q19 Marks", self.__style0)
        ws.write(0, 119, "Actual Q19 Marks", self.__style0)
        ws.write(0, 120, "Expected Q20 Marks", self.__style0)
        ws.write(0, 121, "Actual Q20 Marks", self.__style0)
        ws.write(0, 122, "Expected Q21 Marks", self.__style0)
        ws.write(0, 123, "Actual Q21 Marks", self.__style0)
        ws.write(0, 124, "Expected Q22 Marks", self.__style0)
        ws.write(0, 125, "Actual Q22 Marks", self.__style0)
        ws.write(0, 126, "Expected Q23 Marks", self.__style0)
        ws.write(0, 127, "Actual Q23 Marks", self.__style0)
        ws.write(0, 128, "Expected Q24 Marks", self.__style0)
        ws.write(0, 129, "Actual Q24 Marks", self.__style0)
        ws.write(0, 130, "Expected Q25 Marks", self.__style0)
        ws.write(0, 131, "Actual Q25 Marks", self.__style0)
        ws.write(0, 132, "Expected Q26 Marks", self.__style0)
        ws.write(0, 133, "Actual Q26 Marks", self.__style0)
        ws.write(0, 134, "Expected Q27 Marks", self.__style0)
        ws.write(0, 135, "Actual Q27 Marks", self.__style0)
        ws.write(0, 136, "Expected Q28 Marks", self.__style0)
        ws.write(0, 137, "Actual Q28 Marks", self.__style0)
        ws.write(0, 138, "Expected Q29 Marks", self.__style0)
        ws.write(0, 139, "Actual Q29 Marks", self.__style0)
        ws.write(0, 140, "Expected Q30 Marks", self.__style0)
        ws.write(0, 141, "Actual Q30 Marks", self.__style0)
        ws.write(0, 142, "Expected Q31 Marks", self.__style0)
        ws.write(0, 143, "Actual Q31 Marks", self.__style0)
        ws.write(0, 144, "Expected Q32 Marks", self.__style0)
        ws.write(0, 145, "Actual Q32 Marks", self.__style0)
        ws.write(0, 146, "Expected Q33 Marks", self.__style0)
        ws.write(0, 147, "Actual Q33 Marks", self.__style0)
        ws.write(0, 148, "Expected Q34 Marks", self.__style0)
        ws.write(0, 149, "Actual Q34 Marks", self.__style0)
        ws.write(0, 150, "Expected Q35 Marks", self.__style0)
        ws.write(0, 151, "Actual Q35 Marks", self.__style0)
        ws.write(0, 152, "Expected Q36 Marks", self.__style0)
        ws.write(0, 153, "Actual Q36 Marks", self.__style0)
        ws.write(0, 154, "Expected Q37 Marks", self.__style0)
        ws.write(0, 155, "Actual Q37 Marks", self.__style0)
        ws.write(0, 156, "Expected Q38 Marks", self.__style0)
        ws.write(0, 157, "Actual Q38 Marks", self.__style0)
        ws.write(0, 158, "Expected Q39 Marks", self.__style0)
        ws.write(0, 159, "Actual Q39 Marks", self.__style0)


        n = 1
        rownum = n
        while n < sh1.nrows:
            rows = sh1.row_values(rownum)
            candidateId = int(rows[0])
            loginName = rows[1]
            Passwords = rows[2]
            TestId = int(rows[3])
            SectionId_1 = int(rows[4])
            SectionId_2 = int(rows[7])
            SectionId_3 = int(rows[10])
            SectionId_4 = int(rows[23])
            SectionId_5 = int(rows[28])
            SectionId_6 = int(rows[33])
            SectionId_7 = int(rows[40])
            SectionId_8 = int(rows[44])
            SectionId_9 = int(rows[48])

            QuestionId_1_Ans = rows[5]
            if (QuestionId_1_Ans == True or QuestionId_1_Ans == 1):
                QuestionId_1_Ans = "True"
            elif (QuestionId_1_Ans == False or QuestionId_1_Ans == 0):
                QuestionId_1_Ans = "False"
            else:
                 QuestionId_1_Ans = None
            QuestionId_2_Ans = rows[6]
            if (QuestionId_2_Ans == True or QuestionId_2_Ans == 1):
                QuestionId_2_Ans = "True"
            elif(QuestionId_2_Ans == False or QuestionId_1_Ans == 0):
                QuestionId_2_Ans = "False"
            else:
                 QuestionId_2_Ans = None
                 # print(QuestionId_1_Ans, QuestionId_2_Ans)
            QuestionId_3_Ans = rows[8]
            QuestionId_4_Ans = rows[9]
            QuestionId_5_Ans = rows[11]
            QuestionId_6_Ans = rows[12]
            QuestionId_7_Ans = rows[13]
            QuestionId_8_Ans = rows[14]
            QuestionId_9_Ans = rows[15]
            QuestionId_10_Ans = rows[16]
            QuestionId_11_Ans = rows[17]
            QuestionId_12_Ans = rows[18]
            QuestionId_13_Ans = rows[19]
            QuestionId_14_Ans = rows[20]
            QuestionId_15_Ans = rows[21]
            QuestionId_16_Ans = rows[22]
            QuestionId_17_Ans = rows[24]
            if (QuestionId_17_Ans == True or QuestionId_17_Ans == 1):
                QuestionId_17_Ans = "True"
            elif(QuestionId_17_Ans == False or QuestionId_17_Ans == 0):
                QuestionId_17_Ans = "False"
            else:
                 QuestionId_17_Ans = None

            QuestionId_18_Ans = rows[25]
            if (QuestionId_18_Ans == True or QuestionId_18_Ans == 1):
                QuestionId_18_Ans = "True"
            elif(QuestionId_18_Ans == False or QuestionId_18_Ans == 0):
                QuestionId_18_Ans = "False"
            else:
                 QuestionId_17_Ans = None

            QuestionId_19_Ans = rows[26]
            if (QuestionId_19_Ans == True or QuestionId_19_Ans == 1):
                QuestionId_19_Ans = "True"
            elif(QuestionId_19_Ans == False or QuestionId_19_Ans == 0):
                QuestionId_19_Ans = "False"
            else:
                QuestionId_19_Ans = None

            QuestionId_20_Ans = rows[27]
            if (QuestionId_20_Ans == True or QuestionId_20_Ans == 1):
                QuestionId_20_Ans = "True"
            elif(QuestionId_20_Ans == False or QuestionId_20_Ans == 0):
                QuestionId_20_Ans = "False"
            else:
                QuestionId_20_Ans = None

            # print(QuestionId_17_Ans, QuestionId_18_Ans, QuestionId_19_Ans, QuestionId_20_Ans)
            QuestionId_21_Ans = rows[29]
            QuestionId_22_Ans = rows[30]
            QuestionId_23_Ans = rows[31]
            QuestionId_24_Ans = rows[32]
            QuestionId_25_Ans = rows[34]
            QuestionId_26_Ans = rows[35]
            QuestionId_27_Ans = rows[36]
            QuestionId_28_Ans = rows[37]
            QuestionId_29_Ans = rows[38]
            QuestionId_30_Ans = rows[39]
            QuestionId_31_Ans = rows[41]
            QuestionId_32_Ans = rows[42]
            QuestionId_33_Ans = rows[43]
            QuestionId_34_Ans = rows[45]
            QuestionId_35_Ans = rows[46]
            QuestionId_36_Ans = rows[47]
            QuestionId_37_Ans = rows[49]
            if (QuestionId_37_Ans == True or QuestionId_37_Ans == 1):
                QuestionId_37_Ans = "True"
            elif(QuestionId_37_Ans == False or QuestionId_37_Ans == 0):
                QuestionId_37_Ans = "False"
            else:
                QuestionId_37_Ans = None

            QuestionId_38_Ans = rows[50]
            if (QuestionId_38_Ans == True or QuestionId_38_Ans == 1):
                QuestionId_38_Ans = "True"
            elif(QuestionId_38_Ans == False or QuestionId_38_Ans == 0):
                QuestionId_38_Ans = "False"
            else:
                QuestionId_38_Ans = None

            QuestionId_39_Ans = rows[51]
            if (QuestionId_39_Ans == True or QuestionId_39_Ans == 1):
                QuestionId_39_Ans = "True"
            elif(QuestionId_39_Ans == False or QuestionId_39_Ans == 0):
                QuestionId_39_Ans = "False"
            else:
                QuestionId_39_Ans = None
            # print(QuestionId_37_Ans, QuestionId_38_Ans, QuestionId_39_Ans

            # Expected_Sec_1_Total = rows[52]
            # Expected_Sec_2_Total = rows[54]
            # Expected_Sec_3_Total = rows[56]
            # Expected_Sec_4_Total = rows[58]
            # Expected_Sec_5_Total = rows[60]
            # Expected_Sec_6_Total = rows[62]
            # Expected_Sec_7_Total = rows[64]
            # Expected_Sec_8_Total = rows[66]
            # Expected_Sec_9_Total = rows[68]

            Expected_Grp_1_Total = rows[70]
            Expected_Grp_2_Total = rows[72]
            Expected_Grp_3_Total = rows[74]
            Expected_Grp_4_Total = rows[76]

            Expected_test_Total = rows[78]
            Expected_Percentage = rows[80]

            Expected_Q1_Marks = rows[82]
            Expected_Q2_Marks = rows[84]
            Expected_Q3_Marks = rows[86]
            Expected_Q4_Marks = rows[88]
            Expected_Q5_Marks = rows[90]
            Expected_Q6_Marks = rows[92]
            Expected_Q7_Marks = rows[94]
            Expected_Q8_Marks = rows[96]
            Expected_Q9_Marks = rows[98]
            Expected_Q10_Marks = rows[100]
            Expected_Q11_Marks = rows[102]
            Expected_Q12_Marks = rows[104]
            Expected_Q13_Marks = rows[106]
            Expected_Q14_Marks = rows[108]
            Expected_Q15_Marks = rows[110]
            Expected_Q16_Marks = rows[112]
            Expected_Q17_Marks = rows[114]
            Expected_Q18_Marks = rows[116]
            Expected_Q19_Marks = rows[118]
            Expected_Q20_Marks = rows[120]
            Expected_Q21_Marks = rows[122]
            Expected_Q22_Marks = rows[124]
            Expected_Q23_Marks = rows[126]
            Expected_Q24_Marks = rows[128]
            Expected_Q25_Marks = rows[130]
            Expected_Q26_Marks = rows[132]
            Expected_Q27_Marks = rows[134]
            Expected_Q28_Marks = rows[136]
            Expected_Q29_Marks = rows[138]
            Expected_Q30_Marks = rows[140]
            Expected_Q31_Marks = rows[142]
            Expected_Q32_Marks = rows[144]
            Expected_Q33_Marks = rows[146]
            Expected_Q34_Marks = rows[148]
            Expected_Q35_Marks = rows[150]
            Expected_Q36_Marks = rows[152]
            Expected_Q37_Marks = rows[154]
            Expected_Q38_Marks = rows[156]
            Expected_Q39_Marks = rows[158]


            question_and_answers = {}
            question_and_answers[question_id_headers[0]] = QuestionId_1_Ans
            question_and_answers[question_id_headers[1]] = QuestionId_2_Ans
            question_and_answers[question_id_headers[2]] = QuestionId_3_Ans
            question_and_answers[question_id_headers[3]] = QuestionId_4_Ans
            question_and_answers[question_id_headers[4]] = QuestionId_5_Ans
            question_and_answers[question_id_headers[5]] = QuestionId_6_Ans
            question_and_answers[question_id_headers[6]] = QuestionId_7_Ans
            question_and_answers[question_id_headers[7]] = QuestionId_8_Ans
            question_and_answers[question_id_headers[8]] = QuestionId_9_Ans
            question_and_answers[question_id_headers[9]] = QuestionId_10_Ans
            question_and_answers[question_id_headers[10]] = QuestionId_11_Ans
            question_and_answers[question_id_headers[11]] = QuestionId_12_Ans
            question_and_answers[question_id_headers[12]] = QuestionId_13_Ans
            question_and_answers[question_id_headers[13]] = QuestionId_14_Ans
            question_and_answers[question_id_headers[14]] = QuestionId_15_Ans
            question_and_answers[question_id_headers[15]] = QuestionId_16_Ans
            question_and_answers[question_id_headers[16]] = QuestionId_17_Ans
            question_and_answers[question_id_headers[17]] = QuestionId_18_Ans
            question_and_answers[question_id_headers[18]] = QuestionId_19_Ans
            question_and_answers[question_id_headers[19]] = QuestionId_20_Ans
            question_and_answers[question_id_headers[20]] = QuestionId_21_Ans
            question_and_answers[question_id_headers[21]] = QuestionId_22_Ans
            question_and_answers[question_id_headers[22]] = QuestionId_23_Ans
            question_and_answers[question_id_headers[23]] = QuestionId_24_Ans
            question_and_answers[question_id_headers[24]] = QuestionId_25_Ans
            question_and_answers[question_id_headers[25]] = QuestionId_26_Ans
            question_and_answers[question_id_headers[26]] = QuestionId_27_Ans
            question_and_answers[question_id_headers[27]] = QuestionId_28_Ans
            question_and_answers[question_id_headers[28]] = QuestionId_29_Ans
            question_and_answers[question_id_headers[29]] = QuestionId_30_Ans
            question_and_answers[question_id_headers[30]] = QuestionId_31_Ans
            question_and_answers[question_id_headers[31]] = QuestionId_32_Ans
            question_and_answers[question_id_headers[32]] = QuestionId_33_Ans
            question_and_answers[question_id_headers[33]] = QuestionId_34_Ans
            question_and_answers[question_id_headers[34]] = QuestionId_35_Ans
            question_and_answers[question_id_headers[35]] = QuestionId_36_Ans
            question_and_answers[question_id_headers[36]] = QuestionId_37_Ans
            question_and_answers[question_id_headers[37]] = QuestionId_38_Ans
            question_and_answers[question_id_headers[38]] = QuestionId_39_Ans

            expected_Q_cell_Pos = {}
            expected_Q_cell_Pos[Expected_Q1_Mark] = 82
            expected_Q_cell_Pos[Expected_Q2_Mark] = 84
            expected_Q_cell_Pos[Expected_Q3_Mark] = 86
            expected_Q_cell_Pos[Expected_Q4_Mark] = 88
            expected_Q_cell_Pos[Expected_Q5_Mark] = 90
            expected_Q_cell_Pos[Expected_Q6_Mark] = 92
            expected_Q_cell_Pos[Expected_Q7_Mark] = 94
            expected_Q_cell_Pos[Expected_Q8_Mark] = 96
            expected_Q_cell_Pos[Expected_Q9_Mark] = 98
            expected_Q_cell_Pos[Expected_Q10_Mark] = 100
            expected_Q_cell_Pos[Expected_Q11_Mark] = 102
            expected_Q_cell_Pos[Expected_Q12_Mark] = 104
            expected_Q_cell_Pos[Expected_Q13_Mark] = 106
            expected_Q_cell_Pos[Expected_Q14_Mark] = 108
            expected_Q_cell_Pos[Expected_Q15_Mark] = 110
            expected_Q_cell_Pos[Expected_Q16_Mark] = 112
            expected_Q_cell_Pos[Expected_Q17_Mark] = 114
            expected_Q_cell_Pos[Expected_Q18_Mark] = 116
            expected_Q_cell_Pos[Expected_Q19_Mark] = 118
            expected_Q_cell_Pos[Expected_Q20_Mark] = 120
            expected_Q_cell_Pos[Expected_Q21_Mark] = 122
            expected_Q_cell_Pos[Expected_Q22_Mark] = 124
            expected_Q_cell_Pos[Expected_Q23_Mark] = 126
            expected_Q_cell_Pos[Expected_Q24_Mark] = 128
            expected_Q_cell_Pos[Expected_Q25_Mark] = 130
            expected_Q_cell_Pos[Expected_Q26_Mark] = 132
            expected_Q_cell_Pos[Expected_Q27_Mark] = 134
            expected_Q_cell_Pos[Expected_Q28_Mark] = 136
            expected_Q_cell_Pos[Expected_Q29_Mark] = 138
            expected_Q_cell_Pos[Expected_Q30_Mark] = 140
            expected_Q_cell_Pos[Expected_Q31_Mark] = 142
            expected_Q_cell_Pos[Expected_Q32_Mark] = 144
            expected_Q_cell_Pos[Expected_Q33_Mark] = 146
            expected_Q_cell_Pos[Expected_Q34_Mark] = 148
            expected_Q_cell_Pos[Expected_Q35_Mark] = 150
            expected_Q_cell_Pos[Expected_Q36_Mark] = 152
            expected_Q_cell_Pos[Expected_Q37_Mark] = 154
            expected_Q_cell_Pos[Expected_Q38_Mark] = 156
            expected_Q_cell_Pos[Expected_Q39_Mark] = 158
            print(expected_Q_cell_Pos)

            section_and_expectedmarks = {}
            section_and_expectedmarks[SectionId_1] = 0
            section_and_expectedmarks[SectionId_2] = 0
            section_and_expectedmarks[SectionId_3] = 0
            section_and_expectedmarks[SectionId_4] = 0
            section_and_expectedmarks[SectionId_5] = 0
            section_and_expectedmarks[SectionId_6] = 0
            section_and_expectedmarks[SectionId_7] = 0
            section_and_expectedmarks[SectionId_8] = 0
            section_and_expectedmarks[SectionId_9] = 0
            print("section_and_expectedmarks - ", section_and_expectedmarks)

            group_and_expectedmarks = {}
            # group_and_expectedmarks[Group_1_Id]

            section_expected_cell_pos = {}
            section_expected_cell_pos[Expected_Sec_1_Total_Mark] = 52
            section_expected_cell_pos[Expected_Sec_2_Total_Mark] = 54
            section_expected_cell_pos[Expected_Sec_3_Total_Mark] = 56
            section_expected_cell_pos[Expected_Sec_4_Total_Mark] = 58
            section_expected_cell_pos[Expected_Sec_5_Total_Mark] = 60
            section_expected_cell_pos[Expected_Sec_6_Total_Mark] = 62
            section_expected_cell_pos[Expected_Sec_7_Total_Mark] = 64
            section_expected_cell_pos[Expected_Sec_8_Total_Mark] = 66
            section_expected_cell_pos[Expected_Sec_9_Total_Mark] = 68
            print("section_expected_cell_pos - ", section_expected_cell_pos)



            # ----------------------------------------------------------------------------------------------------------
            # Candidate data printing from Input data excel to Output data Excel
            # ----------------------------------------------------------------------------------------------------------
            ws.write(rownum, 0, candidateId)
            ws.write(rownum, 1, loginName)
            ws.write(rownum, 2, Passwords)
            ws.write(rownum, 3, TestId)
            ws.write(rownum, 4, SectionId_1)
            ws.write(rownum, 5, QuestionId_1_Ans)
            ws.write(rownum, 6, QuestionId_2_Ans)
            ws.write(rownum, 7, SectionId_2)
            ws.write(rownum, 8, QuestionId_3_Ans)
            ws.write(rownum, 9, str(QuestionId_4_Ans))
            ws.write(rownum, 10, SectionId_3)
            ws.write(rownum, 11, QuestionId_5_Ans)
            ws.write(rownum, 12, QuestionId_6_Ans)
            ws.write(rownum, 13, QuestionId_7_Ans)
            ws.write(rownum, 14, QuestionId_8_Ans)
            ws.write(rownum, 15, QuestionId_9_Ans)
            ws.write(rownum, 16, QuestionId_10_Ans)
            ws.write(rownum, 17, QuestionId_11_Ans)
            ws.write(rownum, 18, QuestionId_12_Ans)
            ws.write(rownum, 19, QuestionId_13_Ans)
            ws.write(rownum, 20, QuestionId_14_Ans)
            ws.write(rownum, 21, QuestionId_15_Ans)
            ws.write(rownum, 22, QuestionId_16_Ans)
            ws.write(rownum, 23, SectionId_4)
            ws.write(rownum, 24, QuestionId_17_Ans)
            ws.write(rownum, 25, QuestionId_18_Ans)
            ws.write(rownum, 26, QuestionId_19_Ans)
            ws.write(rownum, 27, QuestionId_20_Ans)
            ws.write(rownum, 28, SectionId_5)
            ws.write(rownum, 29, QuestionId_21_Ans)
            ws.write(rownum, 30, QuestionId_22_Ans)
            ws.write(rownum, 31, QuestionId_23_Ans)
            ws.write(rownum, 32, QuestionId_24_Ans)
            ws.write(rownum, 33, SectionId_6)
            ws.write(rownum, 34, QuestionId_25_Ans)
            ws.write(rownum, 35, QuestionId_26_Ans)
            ws.write(rownum, 36, QuestionId_27_Ans)
            ws.write(rownum, 37, QuestionId_28_Ans)
            ws.write(rownum, 38, QuestionId_29_Ans)
            ws.write(rownum, 39, QuestionId_30_Ans)
            ws.write(rownum, 40, SectionId_7)
            ws.write(rownum, 41, QuestionId_31_Ans)
            ws.write(rownum, 42, QuestionId_32_Ans)
            ws.write(rownum, 43, QuestionId_33_Ans)
            ws.write(rownum, 44, SectionId_8)
            ws.write(rownum, 45, QuestionId_34_Ans)
            ws.write(rownum, 46, QuestionId_35_Ans)
            ws.write(rownum, 47, QuestionId_36_Ans)
            ws.write(rownum, 48, SectionId_9)
            ws.write(rownum, 49, QuestionId_37_Ans)
            ws.write(rownum, 50, QuestionId_38_Ans)
            ws.write(rownum, 51, QuestionId_39_Ans)



            # ws.write(rownum, 14, Expected_Sec_1_Total)
            # ws.write(rownum, 16, Expected_Sec_2_Total)
            # ws.write(rownum, 18, Expected_Sec_3_Total)

            ws.write(rownum, 78, Expected_test_Total)
            ws.write(rownum, 80, Expected_Percentage)
            ws.write(rownum, 82, Expected_Q1_Marks)
            ws.write(rownum, 84, Expected_Q2_Marks)
            ws.write(rownum, 86, Expected_Q3_Marks)
            ws.write(rownum, 88, Expected_Q4_Marks)
            ws.write(rownum, 90, Expected_Q5_Marks)
            ws.write(rownum, 92, Expected_Q6_Marks)
            ws.write(rownum, 94, Expected_Q7_Marks)
            ws.write(rownum, 96, Expected_Q8_Marks)
            ws.write(rownum, 98, Expected_Q9_Marks)
            ws.write(rownum, 100, Expected_Q10_Marks)
            ws.write(rownum, 102, Expected_Q11_Marks)
            ws.write(rownum, 104, Expected_Q12_Marks)
            ws.write(rownum, 106, Expected_Q13_Marks)
            ws.write(rownum, 108, Expected_Q14_Marks)
            ws.write(rownum, 110, Expected_Q15_Marks)
            ws.write(rownum, 112, Expected_Q16_Marks)
            ws.write(rownum, 114, Expected_Q17_Marks)
            ws.write(rownum, 116, Expected_Q18_Marks)
            ws.write(rownum, 118, Expected_Q19_Marks)
            ws.write(rownum, 120, Expected_Q20_Marks)
            ws.write(rownum, 122, Expected_Q21_Marks)
            ws.write(rownum, 124, Expected_Q22_Marks)
            ws.write(rownum, 126, Expected_Q23_Marks)
            ws.write(rownum, 128, Expected_Q24_Marks)
            ws.write(rownum, 130, Expected_Q25_Marks)
            ws.write(rownum, 132, Expected_Q26_Marks)
            ws.write(rownum, 134, Expected_Q27_Marks)
            ws.write(rownum, 136, Expected_Q28_Marks)
            ws.write(rownum, 138, Expected_Q29_Marks)
            ws.write(rownum, 140, Expected_Q30_Marks)
            ws.write(rownum, 142, Expected_Q31_Marks)
            ws.write(rownum, 144, Expected_Q32_Marks)
            ws.write(rownum, 146, Expected_Q33_Marks)
            ws.write(rownum, 148, Expected_Q34_Marks)
            ws.write(rownum, 150, Expected_Q35_Marks)
            ws.write(rownum, 152, Expected_Q36_Marks)
            ws.write(rownum, 154, Expected_Q37_Marks)
            ws.write(rownum, 156, Expected_Q38_Marks)
            ws.write(rownum, 158, Expected_Q39_Marks)


            # ----------------------------------------------------------------------------------------------------------
            # Login to HTML Test/Online Assessment
            # ----------------------------------------------------------------------------------------------------------
            login_to_test_header = {"content-type": "application/json"}
            data1 = {"ClientSystemInfo": "Browser:chrome/60.0.3112.78,OS:Linux x86_64,IPAddress:10.0.3.83",
                     "IPAddress": "10.0.3.83", "IsOnlinePreview": False, "LoginName": loginName,
                     "Password": Passwords,
                     "TenantAlias": "automation"}
            request1 = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v2/login_to_test/",
                                     headers=login_to_test_header, data=json.dumps(data1), verify=True)
            self.Test_Login_response = request1.json()
            self.Test_Login_TokenVal = self.Test_Login_response.get("Token")

            # ----------------------------------------------------------------------------------------------------------
            # Load Test API call
            # ----------------------------------------------------------------------------------------------------------

            loadtest_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
            data2 = {}
            r = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/initiate-tua/", headers=loadtest_header,
                              data=json.dumps(data2, default=str), verify=True)
            dummy = json.loads(r.content)

            data2 = {}
            r = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/loadtest/", headers=loadtest_header,
                              data=json.dumps(data2, default=str), verify=True)
            data = json.loads(r.content)

            questonwise_section = {}
            groups = []
            sections = []
            question_ids = []
            for mandatoryGroups_index in range(0, len(json.loads(r.content)['mandatoryGroups'])):
                groups.append(json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['id'])
                for sections_index in range(0, len(
                        json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'])):
                    sections.append(
                        json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                            'id'])
                    sectionid = \
                        json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                            'id']
                    for questionDetails_index in range(0, len(
                            json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'])):
                        if (json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][sections_index][
                                'questionDetails'][questionDetails_index]['typeOfQuestionText'] == "MCQ"
                            or
                                    json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                        sections_index]['questionDetails'][questionDetails_index][
                                        'typeOfQuestionText'] == "Boolean"):
                            questionid = \
                                json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                    sections_index][
                                    'questionDetails'][questionDetails_index]['id']
                            question_ids.append(
                                json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                    sections_index]['questionDetails'][questionDetails_index]['id'])
                            questonwise_section[questionid] = sectionid
                        else:
                            for childquestionid_index in range(0, len(
                                    json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                        sections_index]['questionDetails'][questionDetails_index]['childQuestions'])):
                                questionid = \
                                    json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                        sections_index]['questionDetails'][questionDetails_index]['childQuestions'][
                                        childquestionid_index]['id']
                                question_ids.append(
                                    json.loads(r.content)['mandatoryGroups'][mandatoryGroups_index]['sections'][
                                        sections_index]['questionDetails'][questionDetails_index]['childQuestions'][
                                        childquestionid_index]['id'])
                                questonwise_section[questionid] = sectionid
            lengthids = len(question_ids)

            print(lengthids)
            print(question_ids)
            ans = []
            print(question_and_answers)
            for i in range(0, lengthids):
                if question_ids[i] in question_and_answers.keys():
                    ans.append(question_and_answers[question_ids[i]])
            print(ans)

            loaded_q_cell_pos = {key: value for key, value in expected_Q_cell_Pos.items() if key in question_ids}
            print(loaded_q_cell_pos)

            other_q_cell_pos = {key: value for key, value in expected_Q_cell_Pos.items() if key not in question_ids}
            print(other_q_cell_pos)

            for key in other_q_cell_pos:
                ws.write(rownum, other_q_cell_pos[key] + 1, "NA")


            for groupid_index in range(0, len(json.loads(r.content)['mandatoryGroups'])):
                for sectionid_index in range(0, len(json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'])):
                    sectionid = json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['id']
                    for questionid_index in range(0, len(json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'])):
                        if (json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'][questionid_index]['typeOfQuestionText']=="MCQ" or json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'][questionid_index]['typeOfQuestionText']=="Boolean"):
                            id = json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'][questionid_index]['id']
                            section_and_expectedmarks[sectionid] = section_and_expectedmarks[sectionid] + rows[expected_Q_cell_Pos[id]]
                        else:
                            for childquestionid in range(0, len(json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'][questionid_index]['childQuestions'])):
                                id = json.loads(r.content)['mandatoryGroups'][groupid_index]['sections'][sectionid_index]['questionDetails'][questionid_index]['childQuestions'][childquestionid]['id']
                                section_and_expectedmarks[sectionid] = section_and_expectedmarks[sectionid] + rows[expected_Q_cell_Pos[id]]
                                # print(section_and_expectedmarks)

            print(section_and_expectedmarks)
            for key in section_and_expectedmarks:
                expectedcell = section_expected_cell_pos[key]
                print(expectedcell)
                print(section_and_expectedmarks[key])
                print(rownum)
                ws.write(rownum, expectedcell, section_and_expectedmarks[key])

            # ----------------------------------------------------------------------------------------------------------
            #  Submit Test API Call
            # ----------------------------------------------------------------------------------------------------------
            candidate_testResultCollection = list()
            for qid in question_ids:
                candidate_testResultCollection.append(
                    {"q": qid, "timeSpent": 1, "secId": questonwise_section[qid], "a": question_and_answers[qid]})

            submit_test_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.Test_Login_TokenVal}
            data4 = {"isPartialSubmission": False, "totalTimeSpent": 39,
                     "testResultCollection": candidate_testResultCollection,
                     "config": "{\"TimeStamp\":\"2018-03-13T07:28:55.933Z\"}"}
            print(data4)

            request4 = requests.post("https://amsin.hirepro.in/py/assessment/htmltest/api/v1/submitTestResult/",
                                     headers=submit_test_header,
                                     data=json.dumps(data4, default=str), verify=True)
            submit_test_response = json.loads(request4.content)

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
            eval_online_assessment_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.NTokenVal,
                                             "APP-NAME": "crpoassessment"}
            data5 = {"testId": TestId, "candidateIds": [candidateId]}
            request5 = requests.post("https://amsin.hirepro.in/py/assessment/eval/api/v1/eval-online-assessment/",
                                     headers=eval_online_assessment_header,
                                     data=json.dumps(data5, default=str), verify=True)
            evaluateTest_response = json.loads(request5.content)
            # print(evaluateTest_response)

            # ----------------------------------------------------------------------------------------------------------
            #  Fetch question wise candidate marks from DB and match with expected
            # ----------------------------------------------------------------------------------------------------------
            try:
                conn = mysql.connector.connect(host='35.154.36.218',
                                               database='appserver_core',
                                               user='hireprouser',
                                               password='tech@123')
                cursor = conn.cursor()
                cursor.execute("select id from Test_users where test_id = 5365 and candidate_id = %d;" % candidateId)
                test_user_id = cursor.fetchone()
                cursor.execute(
                    "select question_id, obtained_marks from test_results where testuser_id = %d;" % test_user_id)
                data = cursor.fetchall()
                print("Type", type(data))
                data_length = len(data)
                # j = 0
                # v = 0
                # col = 29  # it will be depending upon expected cell
                # Expectedcell = 28  # it will be dynamic now for random question paper
                for key in loaded_q_cell_pos:
                    expectedcell = loaded_q_cell_pos[key]
                    col = expectedcell + 1
                    i = 0

                    while (i <= data_length):
                        if (key == data[i][0]):
                            if (rows[expectedcell] == data[i][1]):
                                ws.write(rownum, col, data[i][1], self.__style3)
                            else:
                                ws.write(rownum, col, data[i][1], self.__style2)
                                print("Question -- Candidate_Id - ", candidateId, "Question_Id - ", key, "Expected_Marks - ", \
                                    rows[expectedcell], "Actual_Marks - ", data[i][1])
                            break
                        i += 1
                        # v += 1
                    # j += 1
                    # col += 2
                    # Expectedcell += 2
                conn.commit()
            except:
                print
            finally:
                conn.close()

            # ----------------------------------------------------------------------------------------------------------
            #  View candidate scores by Candidate Id (DotNet API)
            # ----------------------------------------------------------------------------------------------------------
            ViewCandidateScoreByCandidateId_header = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
            data6 = {"TestId": TestId, "CandidateId": candidateId, "TenantId": "ETg6fWphpuw="}  #automation tenant id = "ETg6fWphpuw="
            request6 = requests.post(
                "https://amsin.hirepro.in/amsweb/JSONServices/JSONAssessmentManagementService.svc/ViewCandidateScoreByCandidateId",
                headers=ViewCandidateScoreByCandidateId_header,
                data=json.dumps(data6, default=str), verify=True)
            print("hello",request6.content)
            transcript_response = json.loads(request6.content)['CandidateScore']['TotalCandidateScore']
            print("Test",transcript_response)

            # ----------------------------------------------------------------------------------------------------------
            #  Entering Actual data in excel and compairing Expected and Actual result
            #  Update all ids based on test and Question paper
            # ----------------------------------------------------------------------------------------------------------


            Expected_Grp_1_Total = section_and_expectedmarks[47164] + section_and_expectedmarks[47165] + section_and_expectedmarks[47166]
            Expected_Grp_2_Total = section_and_expectedmarks[47167] + section_and_expectedmarks[47168]
            Expected_Grp_3_Total = section_and_expectedmarks[47169] + section_and_expectedmarks[47170]
            Expected_Grp_4_Total = section_and_expectedmarks[47171] + section_and_expectedmarks[47172]

            ws.write(rownum, 70, Expected_Grp_1_Total)
            ws.write(rownum, 72, Expected_Grp_2_Total)
            ws.write(rownum, 74, Expected_Grp_3_Total)
            ws.write(rownum, 76, Expected_Grp_4_Total)

            Actual_Group1_Score = 0
            Actual_Group2_Score = 0
            Actual_Group3_Score = 0
            Actual_Group4_Score = 0

            Expected_Test_Total = Expected_Grp_1_Total + Expected_Grp_2_Total + Expected_Grp_3_Total + Expected_Grp_4_Total
            ws.write(rownum, 78, Expected_Test_Total)

            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13040):
                    Actual_Group1_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(Expected_Grp_1_Total, 3) == Actual_Group1_Score):
                        ws.write(rownum, 71, Actual_Group1_Score, self.__style3)
                    else:
                        ws.write(rownum, 71, Actual_Group1_Score, self.__style2)
                        Group_1_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidateId, " Group_1_Id - ", Group_1_Id, "Expected Score - ", Expected_Grp_1_Total, "Actual Score - ", Actual_Group1_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47164):
                    Actual_Section1_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47164], 3) == Actual_Section1_Score):
                        ws.write(rownum, 53, Actual_Section1_Score, self.__style3)
                    else:
                        ws.write(rownum, 53, Actual_Section1_Score, self.__style2)
                        Section_1_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_1_Id - ", Section_1_Id, " Expected Score - ", section_and_expectedmarks[47164], " Actual Score - ", Actual_Section1_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47165):
                    Actual_Section2_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47165], 3) == Actual_Section2_Score):
                        ws.write(rownum, 55, Actual_Section2_Score, self.__style3)
                    else:
                        ws.write(rownum, 55, Actual_Section2_Score, self.__style2)
                        Section_2_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Group -- Candidate_Id - ", candidateId, " Section_2_Id - ", Section_2_Id, "Expected Score - ", section_and_expectedmarks[47165], "Actual Score - ", Actual_Section2_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47166):
                    Actual_Section3_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47166], 3) == Actual_Section3_Score):
                        ws.write(rownum, 57, Actual_Section3_Score, self.__style3)
                    else:
                        ws.write(rownum, 57, Actual_Section3_Score, self.__style2)
                        Section_3_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_3_Id - ", Section_3_Id, " Expected Score - ", section_and_expectedmarks[47166], " Actual Score - ", Actual_Section3_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13041):
                    Actual_Group2_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(Expected_Grp_2_Total, 3) == Actual_Group2_Score):
                        ws.write(rownum, 73, Actual_Group2_Score, self.__style3)
                    else:
                        ws.write(rownum, 73, Actual_Group2_Score, self.__style2)
                        Group_2_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidateId, " Group_2_Id - ", Group_2_Id, "Expected Score - ", Expected_Grp_2_Total, "Actual Score - ", Actual_Group2_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47167):
                    Actual_Section4_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47167], 3) == Actual_Section4_Score):
                        ws.write(rownum, 59, Actual_Section4_Score, self.__style3)
                    else:
                        ws.write(rownum, 59, Actual_Section4_Score, self.__style2)
                        Section_4_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_4_Id - ", Section_4_Id, " Expected Score - ", section_and_expectedmarks[47167], " Actual Score - ", Actual_Section4_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47168):
                    Actual_Section5_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47168], 3) == Actual_Section5_Score):
                        ws.write(rownum, 61, Actual_Section5_Score, self.__style3)
                    else:
                        ws.write(rownum, 61, Actual_Section5_Score, self.__style2)
                        Section_5_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_5_Id - ", Section_5_Id, " Expected Score - ", section_and_expectedmarks[47168], " Actual Score - ", Actual_Section5_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13042):
                    Actual_Group3_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(Expected_Grp_3_Total, 3) == Actual_Group3_Score):
                        ws.write(rownum, 75, Actual_Group3_Score, self.__style3)
                    else:
                        ws.write(rownum, 75, Actual_Group3_Score, self.__style2)
                        Group_3_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidateId, " Group_3_Id - ", Group_3_Id, "Expected Score - ", Expected_Grp_3_Total, "Actual Score - ", Actual_Group3_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47169):
                    Actual_Section6_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47169], 3) == Actual_Section6_Score):
                        ws.write(rownum, 63, Actual_Section6_Score, self.__style3)
                    else:
                        ws.write(rownum, 63, Actual_Section6_Score, self.__style2)
                        Section_6_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_6_Id - ", Section_6_Id, " Expected Score - ", section_and_expectedmarks[47169], " Actual Score - ", Actual_Section6_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47170):
                    Actual_Section7_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47170], 3) == Actual_Section7_Score):
                        ws.write(rownum, 65, Actual_Section7_Score, self.__style3)
                    else:
                        ws.write(rownum, 65, Actual_Section7_Score, self.__style2)
                        Section_7_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_7_Id - ", Section_7_Id, " Expected Score - ", section_and_expectedmarks[47170], " Actual Score - ", Actual_Section7_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 13043):
                    Actual_Group4_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if round(Expected_Grp_4_Total, 3) == Actual_Group4_Score:
                        ws.write(rownum, 77, Actual_Group4_Score, self.__style3)
                    else:
                        ws.write(rownum, 77, Actual_Group4_Score, self.__style2)
                        Group_4_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId']
                        print("Group -- Candidate_Id - ", candidateId, " Group_4_Id - ", Group_4_Id, "Expected Score - ", Expected_Grp_4_Total, "Actual Score - ", Actual_Group4_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47171):
                    Actual_Section8_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47171], 3) == Actual_Section8_Score):
                        ws.write(rownum, 67, Actual_Section8_Score, self.__style3)
                    else:
                        ws.write(rownum, 67, Actual_Section8_Score, self.__style2)
                        Section_8_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_8_Id - ", Section_8_Id, " Expected Score - ", section_and_expectedmarks[47171], " Actual Score - ", Actual_Section8_Score)
                    break
                k += 1
            k = 0
            while (k <= 12):
                if (json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k]['GroupId'] == 47172):
                    Actual_Section9_Score = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                        'Score']
                    if (round(section_and_expectedmarks[47172], 3) == Actual_Section9_Score):
                        ws.write(rownum, 69, Actual_Section9_Score, self.__style3)
                    else:
                        ws.write(rownum, 69, Actual_Section9_Score, self.__style2)
                        Section_9_Id = json.loads(request6.content)['CandidateScore']['TotalCandidateScore'][k][
                            'GroupId']
                        print("Section -- Candidate_Id - ", candidateId, " Section_9_Id - ", Section_9_Id, " Expected Score - ", section_and_expectedmarks[47172], " Actual Score - ", Actual_Section9_Score)
                    break
                k += 1
            Expected_Test_Total
            # Actual_Test_Score = Actual_Group1_Score + Actual_Group2_Score + Actual_Group3_Score + Actual_Group4_Score
            Actual_Test_Score = Actual_Group1_Score + Actual_Group2_Score + Actual_Group3_Score + Actual_Group4_Score
            if (round(Expected_Test_Total, 3) == round(Actual_Test_Score, 3)):
                ws.write(rownum, 79, Actual_Test_Score, self.__style3)
            else:
                ws.write(rownum, 79, Actual_Test_Score, self.__style2)
                print("Section -- Candidate_Id - ", candidateId, " Test Id - ", TestId, " Expected Test Score - ", Expected_test_Total, " Actual Test Score - ", Actual_Test_Score)
            wb_result.save("/home/testingteam/hirepro_automation/API-Automation/Output Data/Assessment/staticrandomQP_Evaluation_Check.xls")
            n += 1
            rownum += 1





            # length = len(dummy)
            # dict = []
            # for item in range(0, length):
            #     dummy1 = json.loads(r.content)['mandatoryGroups'][item]['sections']
            #     length1 = len(dummy1)
            #     for item1 in range(0, length1):
            #         dummy2 = json.loads(r.content)['mandatoryGroups'][item]['sections'][item1]['questionDetails']
            #         length2 = len(dummy2)
            #         for item2 in range(0, length2):
            #             dummy3 = json.loads(r.content)['mandatoryGroups'][item]['sections'][item1]['questionDetails'][item2]['typeOfQuestionText']
            #             # dict.append(dummy3)
            #             dict.append(dummy3)
            # print(dict)
            # print("Load Test - ",json.loads(r.content)['mandatoryGroups'][item]['sections'][item1]['questionDetails'][item2]['childQuestions'][1]['typeOfQuestionText'])