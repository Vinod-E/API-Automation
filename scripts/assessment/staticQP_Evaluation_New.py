from __future__ import absolute_import
import requests
import json
import unittest
import mysql
import xlrd
import xlwt
from mysql import connector


class staticQP_Evaluation(unittest.TestCase):
    def test_staticQP_Evaluation(self):

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
        wb = xlrd.open_workbook(
            "/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Input Data/Assessment/staticQP_Evaluation.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet('Evaluation_Check')
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

        question_ids = [
            QuestionId_1, QuestionId_2, QuestionId_3, QuestionId_4, QuestionId_5, QuestionId_6,
            QuestionId_7, QuestionId_8, QuestionId_9, QuestionId_10, QuestionId_11, QuestionId_12,
            QuestionId_13, QuestionId_14, QuestionId_15, QuestionId_16, QuestionId_17, QuestionId_18,
            QuestionId_19, QuestionId_20, QuestionId_21, QuestionId_22, QuestionId_23, QuestionId_24,
            QuestionId_25, QuestionId_26, QuestionId_27, QuestionId_28, QuestionId_29, QuestionId_30,
            QuestionId_31, QuestionId_32, QuestionId_33, QuestionId_34, QuestionId_35, QuestionId_36,
            QuestionId_37, QuestionId_38, QuestionId_39
        ]

        # --------------------------------------------------------------------------------------------------------------
        # Header printing in Output Excel
        # --------------------------------------------------------------------------------------------------------------
        ws.write(0, 0, "Candidate Id", self.__style0)
        ws.write(0, 1, "Login Name", self.__style0)
        ws.write(0, 2, "Password", self.__style0)

        ws.write(0, 3, "Test Id", self.__style0)

        ws.write(0, 4, "Section 1 Id", self.__style0)
        ws.write(0, 5, QuestionId_1, self.__style0), ws.write(0, 6, QuestionId_2, self.__style0)

        ws.write(0, 7, "Section 2 Id", self.__style0)
        ws.write(0, 8, QuestionId_3, self.__style0), ws.write(0, 9, QuestionId_4, self.__style0)

        ws.write(0, 10, "Section 3 Id", self.__style0)
        ws.write(0, 11, QuestionId_5, self.__style0), ws.write(0, 12, QuestionId_6, self.__style0)
        ws.write(0, 13, QuestionId_7, self.__style0), ws.write(0, 14, QuestionId_8, self.__style0)
        ws.write(0, 15, QuestionId_9, self.__style0), ws.write(0, 16, QuestionId_10, self.__style0)
        ws.write(0, 17, QuestionId_11, self.__style0), ws.write(0, 18, QuestionId_12, self.__style0)
        ws.write(0, 19, QuestionId_13, self.__style0), ws.write(0, 20, QuestionId_14, self.__style0)
        ws.write(0, 21, QuestionId_15, self.__style0), ws.write(0, 22, QuestionId_16, self.__style0)

        ws.write(0, 23, "Section 4 Id", self.__style0)
        ws.write(0, 24, QuestionId_17, self.__style0), ws.write(0, 25, QuestionId_18, self.__style0)
        ws.write(0, 26, QuestionId_19, self.__style0), ws.write(0, 27, QuestionId_20, self.__style0)

        ws.write(0, 28, "Section 5 Id", self.__style0)
        ws.write(0, 29, QuestionId_21, self.__style0), ws.write(0, 30, QuestionId_22, self.__style0)
        ws.write(0, 31, QuestionId_23, self.__style0), ws.write(0, 32, QuestionId_24, self.__style0)

        ws.write(0, 33, "Section 6 Id", self.__style0)
        ws.write(0, 34, QuestionId_25, self.__style0), ws.write(0, 35, QuestionId_26, self.__style0)
        ws.write(0, 36, QuestionId_27, self.__style0), ws.write(0, 37, QuestionId_28, self.__style0)
        ws.write(0, 38, QuestionId_29, self.__style0), ws.write(0, 39, QuestionId_30, self.__style0)

        ws.write(0, 40, "Section 7 Id", self.__style0)
        ws.write(0, 41, QuestionId_31, self.__style0), ws.write(0, 42, QuestionId_32, self.__style0)
        ws.write(0, 43, QuestionId_33, self.__style0)

        ws.write(0, 44, "Section 8 Id", self.__style0)
        ws.write(0, 45, QuestionId_34, self.__style0), ws.write(0, 46, QuestionId_35, self.__style0)
        ws.write(0, 47, QuestionId_36, self.__style0)

        ws.write(0, 48, "Section 9 Id", self.__style0)
        ws.write(0, 49, QuestionId_37, self.__style0), ws.write(0, 50, QuestionId_38, self.__style0)
        ws.write(0, 51, QuestionId_39, self.__style0)

        ws.write(0, 52, "Expected_Sec_1_Total", self.__style0), ws.write(0, 53, "Actual_Sec_1_Total", self.__style0)
        ws.write(0, 54, "Expected_Sec_2_Total", self.__style0), ws.write(0, 55, "Actual_Sec_2_Total", self.__style0)
        ws.write(0, 56, "Expected_Sec_3_Total", self.__style0), ws.write(0, 57, "Actual_Sec_3_Total", self.__style0)
        ws.write(0, 58, "Expected_Sec_4_Total", self.__style0), ws.write(0, 59, "Actual_Sec_4_Total", self.__style0)
        ws.write(0, 60, "Expected_Sec_5_Total", self.__style0), ws.write(0, 61, "Actual_Sec_5_Total", self.__style0)
        ws.write(0, 62, "Expected_Sec_6_Total", self.__style0), ws.write(0, 63, "Actual_Sec_6_Total", self.__style0)
        ws.write(0, 64, "Expected_Sec_7_Total", self.__style0), ws.write(0, 65, "Actual_Sec_7_Total", self.__style0)
        ws.write(0, 66, "Expected_Sec_8_Total", self.__style0), ws.write(0, 67, "Actual_Sec_8_Total", self.__style0)
        ws.write(0, 68, "Expected_Sec_9_Total", self.__style0), ws.write(0, 69, "Actual_Sec_9_Total", self.__style0)

        ws.write(0, 70, "Expected_Grp_1_Total", self.__style0), ws.write(0, 71, "Actual_Grp_1_Total", self.__style0)
        ws.write(0, 72, "Expected_Grp_2_Total", self.__style0), ws.write(0, 73, "Actual_Grp_2_Total", self.__style0)
        ws.write(0, 74, "Expected_Grp_3_Total", self.__style0), ws.write(0, 75, "Actual_Grp_3_Total", self.__style0)
        ws.write(0, 76, "Expected_Grp_4_Total", self.__style0), ws.write(0, 77, "Actual_Grp_4_Total", self.__style0)

        ws.write(0, 78, "Expected_test_Total", self.__style0), ws.write(0, 79, "Actual_test_Total", self.__style0)

        ws.write(0, 80, "Expected Percentage", self.__style0), ws.write(0, 81, "X-GUID", self.__style0)

        ws.write(0, 82, "Expected Q1 Marks", self.__style0), ws.write(0, 83, "Actual Q1 Marks", self.__style0)
        ws.write(0, 84, "Expected Q2 Marks", self.__style0), ws.write(0, 85, "Actual Q2 Marks", self.__style0)
        ws.write(0, 86, "Expected Q3 Marks", self.__style0), ws.write(0, 87, "Actual Q3 Marks", self.__style0)
        ws.write(0, 88, "Expected Q4 Marks", self.__style0), ws.write(0, 89, "Actual Q4 Marks", self.__style0)
        ws.write(0, 90, "Expected Q5 Marks", self.__style0), ws.write(0, 91, "Actual Q5 Marks", self.__style0)
        ws.write(0, 92, "Expected Q6 Marks", self.__style0), ws.write(0, 93, "Actual Q6 Marks", self.__style0)
        ws.write(0, 94, "Expected Q7 Marks", self.__style0), ws.write(0, 95, "Actual Q7 Marks", self.__style0)
        ws.write(0, 96, "Expected Q8 Marks", self.__style0), ws.write(0, 97, "Actual Q8 Marks", self.__style0)
        ws.write(0, 98, "Expected Q9 Marks", self.__style0), ws.write(0, 99, "Actual Q9 Marks", self.__style0)
        ws.write(0, 100, "Expected Q10 Marks", self.__style0), ws.write(0, 101, "Actual Q10 Marks", self.__style0)
        ws.write(0, 102, "Expected Q11 Marks", self.__style0), ws.write(0, 103, "Actual Q11 Marks", self.__style0)
        ws.write(0, 104, "Expected Q12 Marks", self.__style0), ws.write(0, 105, "Actual Q12 Marks", self.__style0)
        ws.write(0, 106, "Expected Q13 Marks", self.__style0), ws.write(0, 107, "Actual Q13 Marks", self.__style0)
        ws.write(0, 108, "Expected Q14 Marks", self.__style0), ws.write(0, 109, "Actual Q14 Marks", self.__style0)
        ws.write(0, 110, "Expected Q15 Marks", self.__style0), ws.write(0, 111, "Actual Q15 Marks", self.__style0)
        ws.write(0, 112, "Expected Q16 Marks", self.__style0), ws.write(0, 113, "Actual Q16 Marks", self.__style0)
        ws.write(0, 114, "Expected Q17 Marks", self.__style0), ws.write(0, 115, "Actual Q17 Marks", self.__style0)
        ws.write(0, 116, "Expected Q18 Marks", self.__style0), ws.write(0, 117, "Actual Q18 Marks", self.__style0)
        ws.write(0, 118, "Expected Q19 Marks", self.__style0), ws.write(0, 119, "Actual Q19 Marks", self.__style0)
        ws.write(0, 120, "Expected Q20 Marks", self.__style0), ws.write(0, 121, "Actual Q20 Marks", self.__style0)
        ws.write(0, 122, "Expected Q21 Marks", self.__style0), ws.write(0, 123, "Actual Q21 Marks", self.__style0)
        ws.write(0, 124, "Expected Q22 Marks", self.__style0), ws.write(0, 125, "Actual Q22 Marks", self.__style0)
        ws.write(0, 126, "Expected Q23 Marks", self.__style0), ws.write(0, 127, "Actual Q23 Marks", self.__style0)
        ws.write(0, 128, "Expected Q24 Marks", self.__style0), ws.write(0, 129, "Actual Q24 Marks", self.__style0)
        ws.write(0, 130, "Expected Q25 Marks", self.__style0), ws.write(0, 131, "Actual Q25 Marks", self.__style0)
        ws.write(0, 132, "Expected Q26 Marks", self.__style0), ws.write(0, 133, "Actual Q26 Marks", self.__style0)
        ws.write(0, 134, "Expected Q27 Marks", self.__style0), ws.write(0, 135, "Actual Q27 Marks", self.__style0)
        ws.write(0, 136, "Expected Q28 Marks", self.__style0), ws.write(0, 137, "Actual Q28 Marks", self.__style0)
        ws.write(0, 138, "Expected Q29 Marks", self.__style0), ws.write(0, 139, "Actual Q29 Marks", self.__style0)
        ws.write(0, 140, "Expected Q30 Marks", self.__style0), ws.write(0, 141, "Actual Q30 Marks", self.__style0)
        ws.write(0, 142, "Expected Q31 Marks", self.__style0), ws.write(0, 143, "Actual Q31 Marks", self.__style0)
        ws.write(0, 144, "Expected Q32 Marks", self.__style0), ws.write(0, 145, "Actual Q32 Marks", self.__style0)
        ws.write(0, 146, "Expected Q33 Marks", self.__style0), ws.write(0, 147, "Actual Q33 Marks", self.__style0)
        ws.write(0, 148, "Expected Q34 Marks", self.__style0), ws.write(0, 149, "Actual Q34 Marks", self.__style0)
        ws.write(0, 150, "Expected Q35 Marks", self.__style0), ws.write(0, 151, "Actual Q35 Marks", self.__style0)
        ws.write(0, 152, "Expected Q36 Marks", self.__style0), ws.write(0, 153, "Actual Q36 Marks", self.__style0)
        ws.write(0, 154, "Expected Q37 Marks", self.__style0), ws.write(0, 155, "Actual Q37 Marks", self.__style0)
        ws.write(0, 156, "Expected Q38 Marks", self.__style0), ws.write(0, 157, "Actual Q38 Marks", self.__style0)
        ws.write(0, 158, "Expected Q39 Marks", self.__style0), ws.write(0, 159, "Actual Q39 Marks", self.__style0)

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

            answer_index_list = [5, 6, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 26, 27, 29, 30, 31,
                                 32, 34, 35, 36, 37, 38, 39, 41, 42, 43, 45, 46, 47, 49, 50, 51]
            answer_list = []
            for item in answer_index_list:
                answer_list.append(self.fetch_answer(rows, item))

            Expected_Sec_1_Total = rows[52]
            Expected_Sec_2_Total = rows[54]
            Expected_Sec_3_Total = rows[56]
            Expected_Sec_4_Total = rows[58]
            Expected_Sec_5_Total = rows[60]
            Expected_Sec_6_Total = rows[62]
            Expected_Sec_7_Total = rows[64]
            Expected_Sec_8_Total = rows[66]
            Expected_Sec_9_Total = rows[68]

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

            # ----------------------------------------------------------------------------------------------------------
            # Candidate data printing from Input data excel to Output data Excel
            # ----------------------------------------------------------------------------------------------------------
            ws.write(rownum, 0, candidateId)
            ws.write(rownum, 1, loginName)
            ws.write(rownum, 2, Passwords)
            ws.write(rownum, 3, TestId)

            ws.write(rownum, 4, SectionId_1)
            ws.write(rownum, 5, answer_list[0]), ws.write(rownum, 6, answer_list[1])

            ws.write(rownum, 7, SectionId_2)
            ws.write(rownum, 8, answer_list[2]), ws.write(rownum, 9, answer_list[3])

            ws.write(rownum, 10, SectionId_3)
            ws.write(rownum, 11, answer_list[4]), ws.write(rownum, 12, answer_list[5])
            ws.write(rownum, 13, answer_list[6]), ws.write(rownum, 14, answer_list[7])
            ws.write(rownum, 15, answer_list[8]), ws.write(rownum, 16, answer_list[9])
            ws.write(rownum, 17, answer_list[10]), ws.write(rownum, 18, answer_list[11])
            ws.write(rownum, 19, answer_list[12]), ws.write(rownum, 20, answer_list[13])
            ws.write(rownum, 21, answer_list[14]), ws.write(rownum, 22, answer_list[15])

            ws.write(rownum, 23, SectionId_4)
            ws.write(rownum, 24, answer_list[16]), ws.write(rownum, 25, answer_list[17])
            ws.write(rownum, 26, answer_list[18]), ws.write(rownum, 27, answer_list[19])

            ws.write(rownum, 28, SectionId_5)
            ws.write(rownum, 29, answer_list[20]), ws.write(rownum, 30, answer_list[21])
            ws.write(rownum, 31, answer_list[22]), ws.write(rownum, 32, answer_list[23])

            ws.write(rownum, 33, SectionId_6)
            ws.write(rownum, 34, answer_list[24]), ws.write(rownum, 35, answer_list[25])
            ws.write(rownum, 36, answer_list[26]), ws.write(rownum, 37, answer_list[27])
            ws.write(rownum, 38, answer_list[28]), ws.write(rownum, 39, answer_list[29])

            ws.write(rownum, 40, SectionId_7)
            ws.write(rownum, 41, answer_list[30]), ws.write(rownum, 42, answer_list[31])
            ws.write(rownum, 43, answer_list[32])

            ws.write(rownum, 44, SectionId_8)
            ws.write(rownum, 45, answer_list[33]), ws.write(rownum, 46, answer_list[34])
            ws.write(rownum, 47, answer_list[35])

            ws.write(rownum, 48, SectionId_9)
            ws.write(rownum, 49, answer_list[36]), ws.write(rownum, 50, answer_list[37])
            ws.write(rownum, 51, answer_list[38])

            ws.write(rownum, 52, Expected_Sec_1_Total)
            ws.write(rownum, 54, Expected_Sec_2_Total)
            ws.write(rownum, 56, Expected_Sec_3_Total)
            ws.write(rownum, 58, Expected_Sec_4_Total)
            ws.write(rownum, 60, Expected_Sec_5_Total)
            ws.write(rownum, 62, Expected_Sec_6_Total)
            ws.write(rownum, 64, Expected_Sec_7_Total)
            ws.write(rownum, 66, Expected_Sec_8_Total)
            ws.write(rownum, 68, Expected_Sec_9_Total)

            ws.write(rownum, 70, Expected_Grp_1_Total)
            ws.write(rownum, 72, Expected_Grp_2_Total)
            ws.write(rownum, 74, Expected_Grp_3_Total)
            ws.write(rownum, 76, Expected_Grp_4_Total)

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
            login_to_test_header = {"content-type": "application/json", "X-APPLMA": "true"}
            login_to_test_data = {"ClientSystemInfo": "Browser:chrome/60.0.3112.78,OS:Linux x86_64,IPAddress:10.0.3.83",
                                  "IPAddress": "10.0.3.83", "IsOnlinePreview": False, "LoginName": loginName,
                                  "Password": Passwords,
                                  "TenantAlias": "automation"}
            login_to_test_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/htmltest/api/v2/login_to_test/",
                headers=login_to_test_header, data=json.dumps(login_to_test_data), verify=True)
            self.login_to_test_response = login_to_test_request.json()
            print(self.login_to_test_response)
            self.login_to_test_TokenVal = self.login_to_test_response.get("Token")

            # ----------------------------------------------------------------------------------------------------------
            #  Submit Test API Call
            # ----------------------------------------------------------------------------------------------------------
            submitTestResult_header = {"content-type": "application/json", "X-APPLMA": "true",
                                       "X-AUTH-TOKEN": self.login_to_test_TokenVal}
            submitTestResult_data = {"isPartialSubmission": False, "totalTimeSpent": 39,
                                     "testResultCollection": [
                                         {"q": QuestionId_1, "timeSpent": 1, "secId": SectionId_1, "a": answer_list[0]},
                                         {"q": QuestionId_2, "timeSpent": 1, "secId": SectionId_1, "a": answer_list[1]},
                                         {"q": QuestionId_3, "timeSpent": 1, "secId": SectionId_2, "a": answer_list[2]},
                                         {"q": QuestionId_4, "timeSpent": 1, "secId": SectionId_2, "a": answer_list[3]},
                                         {"q": QuestionId_5, "timeSpent": 1, "secId": SectionId_3, "a": answer_list[4]},
                                         {"q": QuestionId_6, "timeSpent": 1, "secId": SectionId_3, "a": answer_list[5]},
                                         {"q": QuestionId_7, "timeSpent": 1, "secId": SectionId_3, "a": answer_list[6]},
                                         {"q": QuestionId_8, "timeSpent": 1, "secId": SectionId_3, "a": answer_list[7]},
                                         {"q": QuestionId_9, "timeSpent": 1, "secId": SectionId_3, "a": answer_list[8]},
                                         {"q": QuestionId_10, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[9]},
                                         {"q": QuestionId_11, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[10]},
                                         {"q": QuestionId_12, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[11]},
                                         {"q": QuestionId_13, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[12]},
                                         {"q": QuestionId_14, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[13]},
                                         {"q": QuestionId_15, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[14]},
                                         {"q": QuestionId_16, "timeSpent": 1, "secId": SectionId_3,
                                          "a": answer_list[15]},
                                         {"q": QuestionId_17, "timeSpent": 1, "secId": SectionId_4,
                                          "a": answer_list[16]},
                                         {"q": QuestionId_18, "timeSpent": 1, "secId": SectionId_4,
                                          "a": answer_list[17]},
                                         {"q": QuestionId_19, "timeSpent": 1, "secId": SectionId_4,
                                          "a": answer_list[18]},
                                         {"q": QuestionId_20, "timeSpent": 1, "secId": SectionId_4,
                                          "a": answer_list[19]},
                                         {"q": QuestionId_21, "timeSpent": 1, "secId": SectionId_5,
                                          "a": answer_list[20]},
                                         {"q": QuestionId_22, "timeSpent": 1, "secId": SectionId_5,
                                          "a": answer_list[21]},
                                         {"q": QuestionId_23, "timeSpent": 1, "secId": SectionId_5,
                                          "a": answer_list[22]},
                                         {"q": QuestionId_24, "timeSpent": 1, "secId": SectionId_5,
                                          "a": answer_list[23]},
                                         {"q": QuestionId_25, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[24]},
                                         {"q": QuestionId_26, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[25]},
                                         {"q": QuestionId_27, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[26]},
                                         {"q": QuestionId_28, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[27]},
                                         {"q": QuestionId_29, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[28]},
                                         {"q": QuestionId_30, "timeSpent": 1, "secId": SectionId_6,
                                          "a": answer_list[29]},
                                         {"q": QuestionId_31, "timeSpent": 1, "secId": SectionId_7,
                                          "a": answer_list[30]},
                                         {"q": QuestionId_32, "timeSpent": 1, "secId": SectionId_7,
                                          "a": answer_list[31]},
                                         {"q": QuestionId_33, "timeSpent": 1, "secId": SectionId_7,
                                          "a": answer_list[32]},
                                         {"q": QuestionId_34, "timeSpent": 1, "secId": SectionId_8,
                                          "a": answer_list[33]},
                                         {"q": QuestionId_35, "timeSpent": 1, "secId": SectionId_8,
                                          "a": answer_list[34]},
                                         {"q": QuestionId_36, "timeSpent": 1, "secId": SectionId_8,
                                          "a": answer_list[35]},
                                         {"q": QuestionId_37, "timeSpent": 1, "secId": SectionId_9,
                                          "a": answer_list[36]},
                                         {"q": QuestionId_38, "timeSpent": 1, "secId": SectionId_9,
                                          "a": answer_list[37]},
                                         {"q": QuestionId_39, "timeSpent": 1, "secId": SectionId_9,
                                          "a": answer_list[38]}
                                     ],
                                     "config": "{\"TimeStamp\":\"2018-03-27T07:28:55.933Z\"}"}

            submitTestResult_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/htmltest/api/v1/submitTestResult/",
                headers=submitTestResult_header,
                data=json.dumps(submitTestResult_data, default=str), verify=True)
            submitTestResult_response = json.loads(submitTestResult_request.content)
            print(submitTestResult_response)

            # ----------------------------------------------------------------------------------------------------------
            #  Login to AMS
            # ----------------------------------------------------------------------------------------------------------
            login_user_header = {"content-type": "application/json", "X-APPLMA": "true"}
            login_user_data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation",
                               "UserName": "admin"}
            login_user_response = requests.post('https://amsin.hirepro.in/py/common/user/login_user/',
                                                headers=login_user_header, data=json.dumps(login_user_data),
                                                verify=True)
            self.login_user_TokenVal = login_user_response.json()

            # ----------------------------------------------------------------------------------------------------------
            #  Evaluate online assessment for candidate
            # ----------------------------------------------------------------------------------------------------------
            eval_online_assessment_header = {"content-type": "application/json",
                                             "X-AUTH-TOKEN": self.login_user_TokenVal.get("Token"),
                                             "APP-NAME": "crpoassessment"}
            eval_online_assessment_data = {"testId": TestId, "candidateIds": [candidateId]}
            eval_online_assessment_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/eval/api/v1/eval-online-assessment/",
                headers=eval_online_assessment_header,
                data=json.dumps(eval_online_assessment_data, default=str), verify=True)
            eval_online_assessment_response = json.loads(eval_online_assessment_request.content)
            GUID = eval_online_assessment_request.headers['X-GUID']
            ws.write(rownum, 81, GUID, self.__style4)

            # ----------------------------------------------------------------------------------------------------------
            #  Fetch question wise candidate marks from DB and match with expected
            # ----------------------------------------------------------------------------------------------------------
            try:
                conn = mysql.connector.connect(host='35.154.36.218',
                                               database='appserver_core',
                                               user='hireprouser',
                                               password='tech@123')
                cursor = conn.cursor()
                cursor.execute("select id from Test_users where test_id = 5282 and candidate_id = %d;" % candidateId)
                test_user_id = cursor.fetchone()
                cursor.execute(
                    "select question_id, obtained_marks from test_results where testuser_id = %d;" % test_user_id)
                data = cursor.fetchall()
                print("Variable Type", type(data))
                print("My Data ", data)
                data_length = len(data)
                j = 0

                col = 83
                Expectedcell = 82
                while (j <= len(question_ids)):
                    i = 0
                    while (i <= data_length):
                        if (question_ids[j] == data[i][0]):
                            if (rows[Expectedcell] == data[i][1]):
                                ws.write(rownum, col, data[i][1], self.__style3)
                            else:
                                ws.write(rownum, col, data[i][1], self.__style2)
                                print("Question -- Candidate_Id - ", candidateId, "Question_Id - ", question_ids[j],
                                      "Expected_Marks - ", \
                                      rows[Expectedcell], "Actual_Marks - ", data[i][1])
                            break
                        i += 1
                    j += 1
                    col += 2
                    Expectedcell += 2
                conn.commit()
            except:
                print
            finally:
                conn.close()

            # ----------------------------------------------------------------------------------------------------------
            #  View candidate scores by Candidate Id
            # ----------------------------------------------------------------------------------------------------------
            view_candidate_score_by_candidate_id_header = {"content-type": "application/json",
                                                           "X-AUTH-TOKEN": self.login_user_TokenVal.get("Token")}
            view_candidate_score_by_candidate_id_data = {"testId": TestId, "candidateId": candidateId,
                                                         "reportFlags": {'testUsersScoreRequired': True,
                                                                         'fileContentRequired': False}, "print": False}

            print(view_candidate_score_by_candidate_id_data)
            view_candidate_score_by_candidate_id_request = requests.post(
                "https://amsin.hirepro.in/py/assessment/report/api/v1/candidatetranscript/",
                headers=view_candidate_score_by_candidate_id_header,
                data=json.dumps(view_candidate_score_by_candidate_id_data, default=str), verify=True)
            print(view_candidate_score_by_candidate_id_request)

            # ----------------------------------------------------------------------------------------------------------
            #  Entering Actual data in excel and compairing Expected and Actual result
            # ----------------------------------------------------------------------------------------------------------
            Actual_Group1_Score = 0
            Actual_Group2_Score = 0
            Actual_Group3_Score = 0
            Actual_Group4_Score = 0

            testUserScore = len(
                json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'])
            for ite in range(0, testUserScore):
                if json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite][
                    'candidateId'] == candidateId:
                    k = 0
                    while k <= 13:
                        if \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite][
                            'groupInfos'][0]['groupId'] == 12857:
                            Actual_Group1_Score = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][
                                ite]['groupInfos'][0]['score']
                            if round(Expected_Grp_1_Total, 3) == round(Actual_Group1_Score, 3):
                                ws.write(rownum, 71, Actual_Group1_Score, self.__style3)
                            else:
                                ws.write(rownum, 71, Actual_Group1_Score, self.__style2)
                                group_1_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][0]['groupId']
                                print("Group -- Candidate_Id - ", candidateId, " group_1_id - ", group_1_id,
                                      "Expected Score - ", Expected_Grp_1_Total, "Actual Score - ", Actual_Group1_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite][
                            'groupInfos'][0]['sectionInfos'][k]["sectionId"] == 46869:
                            Actual_Section1_Score = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][
                                ite]['groupInfos'][0]['sectionInfos'][k]['score']
                            if round(Expected_Sec_1_Total, 3) == round(Actual_Section1_Score, 3):
                                ws.write(rownum, 53, Actual_Section1_Score, self.__style3)
                            else:
                                ws.write(rownum, 53, Actual_Section1_Score, self.__style2)
                                section_1_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_1_id,
                                      "Expected Score - ", Expected_Sec_1_Total, " Actual Score - ",
                                      Actual_Section1_Score)
                            break
                        k += 1

                    k = 0
                    while k <= 13:
                        if \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite][
                            'groupInfos'][0]['sectionInfos'][k]["sectionId"] == 46870:
                            Actual_Section2_Score = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][
                                ite]['groupInfos'][0]['sectionInfos'][k]['score']
                            if round(Expected_Sec_2_Total, 3) == round(Actual_Section2_Score, 3):
                                ws.write(rownum, 55, Actual_Section2_Score, self.__style3)
                            else:
                                ws.write(rownum, 55, Actual_Section2_Score, self.__style2)
                                section_2_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_2_id,
                                      "Expected Score - ", Expected_Sec_2_Total, " Actual Score - ",
                                      Actual_Section2_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                        json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][ite][
                            'groupInfos'][0]['sectionInfos'][k]["sectionId"] == 46871:
                            Actual_Section3_Score = \
                            json.loads(view_candidate_score_by_candidate_id_request.content)['data']['testUserScore'][
                                ite]['groupInfos'][0]['sectionInfos'][k]['score']
                            if round(Expected_Sec_3_Total, 3) == round(Actual_Section3_Score, 3):
                                ws.write(rownum, 57, Actual_Section3_Score, self.__style3)
                            else:
                                ws.write(rownum, 57, Actual_Section3_Score, self.__style2)
                                section_3_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][0]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_3_id,
                                      "Expected Score - ", Expected_Sec_3_Total, " Actual Score - ",
                                      Actual_Section3_Score)
                            break
                        k += 1

                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][1]['groupId'] == 12858:
                            Actual_Group2_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][1]['score']
                            if round(Expected_Grp_2_Total, 3) == round(Actual_Group2_Score, 3):
                                ws.write(rownum, 73, Actual_Group2_Score, self.__style3)
                            else:
                                ws.write(rownum, 73, Actual_Group2_Score, self.__style2)
                                group_2_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][1]['groupId']
                                print("Group -- Candidate_Id - ", candidateId, " group_1_id - ", group_2_id,
                                      "Expected Score - ", Expected_Grp_2_Total, "Actual Score - ", Actual_Group2_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][1]['sectionInfos'][k]["sectionId"] == 46872:
                            Actual_Section4_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][1]['sectionInfos'][k]['score']
                            if round(Expected_Sec_4_Total, 3) == round(Actual_Section4_Score, 3):
                                ws.write(rownum, 59, Actual_Section4_Score, self.__style3)
                            else:
                                ws.write(rownum, 59, Actual_Section4_Score, self.__style2)
                                section_4_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_4_id,
                                      "Expected Score - ", Expected_Sec_4_Total, " Actual Score - ",
                                      Actual_Section4_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][1]['sectionInfos'][k]["sectionId"] == 46873:
                            Actual_Section5_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][1]['sectionInfos'][k]['score']
                            if round(Expected_Sec_5_Total, 3) == round(Actual_Section5_Score, 3):
                                ws.write(rownum, 61, Actual_Section5_Score, self.__style3)
                            else:
                                ws.write(rownum, 61, Actual_Section5_Score, self.__style2)
                                section_5_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][1]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_5_id,
                                      "Expected Score - ", Expected_Sec_5_Total, " Actual Score - ",
                                      Actual_Section5_Score)
                            break
                        k += 1

                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][2]['groupId'] == 12859:
                            Actual_Group3_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][2]['score']
                            if round(Expected_Grp_3_Total, 3) == round(Actual_Group3_Score, 3):
                                ws.write(rownum, 75, Actual_Group3_Score, self.__style3)
                            else:
                                ws.write(rownum, 75, Actual_Group3_Score, self.__style2)
                                group_3_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][2]['groupId']
                                print("Group -- Candidate_Id - ", candidateId, " group_1_id - ", group_3_id,
                                      "Expected Score - ", Expected_Grp_3_Total, "Actual Score - ", Actual_Group3_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][2]['sectionInfos'][k]["sectionId"] == 46874:
                            Actual_Section6_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][2]['sectionInfos'][k]['score']
                            if round(Expected_Sec_6_Total, 3) == round(Actual_Section6_Score, 3):
                                ws.write(rownum, 63, Actual_Section6_Score, self.__style3)
                            else:
                                ws.write(rownum, 63, Actual_Section6_Score, self.__style2)
                                section_6_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][2]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_6_id,
                                      "Expected Score - ", Expected_Sec_6_Total, " Actual Score - ",
                                      Actual_Section6_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][2]['sectionInfos'][k]["sectionId"] == 46875:
                            Actual_Section7_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][2]['sectionInfos'][k]['score']
                            if round(Expected_Sec_7_Total, 3) == round(Actual_Section7_Score, 3):
                                ws.write(rownum, 65, Actual_Section7_Score, self.__style3)
                            else:
                                ws.write(rownum, 65, Actual_Section7_Score, self.__style2)
                                section_7_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][2]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_7_id,
                                      "Expected Score - ", Expected_Sec_7_Total, " Actual Score - ",
                                      Actual_Section7_Score)
                            break
                        k += 1

                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][3]['groupId'] == 12860:
                            Actual_Group4_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][3]['score']
                            if round(Expected_Grp_4_Total, 3) == round(Actual_Group4_Score, 3):
                                ws.write(rownum, 77, Actual_Group4_Score, self.__style3)
                            else:
                                ws.write(rownum, 77, Actual_Group4_Score, self.__style2)
                                group_4_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][3]['groupId']
                                print("Group -- Candidate_Id - ", candidateId, " group_1_id - ", group_4_id,
                                      "Expected Score - ", Expected_Grp_4_Total, "Actual Score - ", Actual_Group4_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][3]['sectionInfos'][k]["sectionId"] == 46876:
                            Actual_Section8_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][3]['sectionInfos'][k]['score']
                            if round(Expected_Sec_8_Total, 3) == round(Actual_Section8_Score, 3):
                                ws.write(rownum, 67, Actual_Section8_Score, self.__style3)
                            else:
                                ws.write(rownum, 67, Actual_Section8_Score, self.__style2)
                                section_8_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][3]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_8_id,
                                      "Expected Score - ", Expected_Sec_8_Total, " Actual Score - ",
                                      Actual_Section8_Score)
                            break
                        k += 1
                    k = 0
                    while k <= 13:
                        if \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite][
                                    'groupInfos'][3]['sectionInfos'][k]["sectionId"] == 46877:
                            Actual_Section9_Score = \
                                json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][
                                    ite]['groupInfos'][3]['sectionInfos'][k]['score']
                            if round(Expected_Sec_9_Total, 3) == round(Actual_Section9_Score, 3):
                                ws.write(rownum, 69, Actual_Section9_Score, self.__style3)
                            else:
                                ws.write(rownum, 69, Actual_Section9_Score, self.__style2)
                                section_9_id = json.loads(view_candidate_score_by_candidate_id_request.content)['data'][
                                    'testUserScore'][ite]['groupInfos'][3]['sectionInfos'][k]["sectionId"]
                                print("Section -- Candidate_Id - ", candidateId, " section_1_id - ", section_9_id,
                                      "Expected Score - ", Expected_Sec_9_Total, " Actual Score - ",
                                      Actual_Section9_Score)
                            break
                        k += 1

            Actual_Test_Score = Actual_Group1_Score + Actual_Group2_Score + Actual_Group3_Score + Actual_Group4_Score
            if (Expected_test_Total == Actual_Test_Score):
                ws.write(rownum, 79, Actual_Test_Score, self.__style3)
            else:
                ws.write(rownum, 79, Actual_Test_Score, self.__style2)
                print("Section -- Candidate_Id - ", candidateId, " Test Id - ", TestId, " Expected Test Score - ",
                      Expected_test_Total, " Actual Test Score - ", Actual_Test_Score)
            wb_result.save(
                "/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/StaticQP_Evaluation_Check.xls")
            n += 1
            rownum += 1

    def fetch_answer(self, rows, cell_index):
        QuestionId_Ans = rows[cell_index]
        if (QuestionId_Ans == True):
            QuestionId_Ans = "True"
        elif (QuestionId_Ans == False):
            QuestionId_Ans = "False"
        elif (QuestionId_Ans == "NA"):
            QuestionId_Ans = None
        return QuestionId_Ans
