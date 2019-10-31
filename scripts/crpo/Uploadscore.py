import time
from hpro_automation import (login, input_paths, output_paths, work_book)
import json
import requests
import xlrd
import datetime


class UploadScoresheet(login.CommonLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(UploadScoresheet, self).__init__()
        self.common_login('crpo')

        self.rowsize1 = 18
        self.size1 = self.rowsize1

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_candidateId = []  # [] Initialising data from excel sheet to the variables
        self.xl_testId = []
        self.xl_group1 = []
        self.xl_section1 = []
        self.xl_section1_1 = []
        self.xl_section1_2 = []
        self.xl_group2 = []
        self.xl_section2 = []
        self.xl_section2_1 = []
        self.xl_group3 = []
        self.xl_section3 = []
        self.xl_section3_1 = []
        self.xl_section2_1 = []
        self.xl_group4 = []
        self.xl_section4 = []
        self.xl_section4_1 = []

        # -------------------------------------
        # Initialising group_section_excel_data
        # -------------------------------------
        self.xl_s1 = []
        self.xl_s2 = []
        self.xl_s3 = []
        self.xl_s4 = []
        self.xl_s5 = []
        self.xl_s6 = []
        self.xl_s7 = []
        self.xl_s8 = []
        self.xl_s9 = []

        # ---------------------------------------------
        # Initialising updated_group_section_excel_data
        # ---------------------------------------------
        self.xl_s1_updated = []
        self.xl_s2_updated = []
        self.xl_s3_updated = []
        self.xl_s4_updated = []
        self.xl_s5_updated = []
        self.xl_s6_updated = []
        self.xl_s7_updated = []
        self.xl_s8_updated = []
        self.xl_s9_updated = []

        # -----------------------------
        # Initialising Group_excel_data
        # -----------------------------
        self.xl_g1 = []
        self.xl_g2 = []
        self.xl_g3 = []
        self.xl_g4 = []
        self.xl_TotalMarks = []

        # -------------------------------------
        # Initialising Updated_Group_excel_data
        # -------------------------------------
        self.xl_g1_updated = []
        self.xl_g2_updated = []
        self.xl_g3_updated = []
        self.xl_g4_updated = []
        self.xl_TotalMarks_updated = []

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.testuser_dict = {}
        self.test_details_dict = {}
        self.candidate_info_dict = {}

        self.group1_dict = {}
        self.section1_dict = {}
        self.section1_1_dict = {}
        self.section1_2_dict = {}

        self.group2_dict = {}
        self.section2_dict = {}
        self.section2_1_dict = {}

        self.group3_dict = {}
        self.section3_dict = {}
        self.section3_1_dict = {}

        self.group4_dict = {}
        self.section4_dict = {}
        self.section4_1_dict = {}

    def excel_headers(self):
        self.main_headers = ['Comparision', 'Candidate Id', 'Overall_Status', 'CandidateName', 'Email', 'Test Mode',
                             'TestStatus', 'TotalMarks', 'Group1', 'S1', 'S2', 'S3', 'Group2', 'S4', 'S5', 'Group3',
                             'S6', 'S7', 'Group4', 'S8', 'S9']
        self.headers_with_style2 = ['Comparision', 'Candidate Id', 'CandidateName', 'Email', 'Mobile', 'Overall_Status']
        self.headers_with_style9 = ['Test Mode', 'TotalMarks', 'TestStatus']
        self.headers_with_style19 = ['Group1', 'Group2', 'Group3', 'Group4']
        self.file_headers_col_row()

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Uploadscore_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)

                self.xl_candidateId.append(int(rows[0]))
                self.xl_testId.append(int(rows[1]))
                self.xl_group1.append(int(rows[3]))
                self.xl_section1.append(int(rows[4]))
                self.xl_section1_1.append(int(rows[5]))
                self.xl_section1_2.append(int(rows[6]))
                self.xl_group2.append(int(rows[7]))
                self.xl_section2.append(int(rows[8]))
                self.xl_section2_1.append(int(rows[9]))
                self.xl_group3.append(int(rows[10]))
                self.xl_section3.append(int(rows[11]))
                self.xl_section3_1.append(int(rows[12]))
                self.xl_group4.append(int(rows[13]))
                self.xl_section4.append(int(rows[14]))
                self.xl_section4_1.append(int(rows[15]))
        except IOError:
            print("File not found or path is incorrect")

    def group_section_excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Group_Section_Input_sheet'])
            sheet2 = workbook.sheet_by_index(0)
            for i in range(1, sheet2.nrows):
                number = i  # Counting number of rows
                rows = sheet2.row_values(number)

                # self.xl_GsCandidate.append(int(rows[0]))

                if rows[9] is not None and rows[9] != '':
                    self.xl_s1.append(int(rows[9]))
                else:
                    self.xl_s1.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_s2.append(int(rows[10]))
                else:
                    self.xl_s2.append(None)

                if rows[11] is not None and rows[11] != '':
                    self.xl_s3.append(int(rows[11]))
                else:
                    self.xl_s3.append(None)

                if rows[12] is not None and rows[12] != '':
                    self.xl_s4.append(int(rows[12]))
                else:
                    self.xl_s4.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_s5.append(int(rows[13]))
                else:
                    self.xl_s5.append(None)

                if rows[14] is not None and rows[14] != '':
                    self.xl_s6.append(int(rows[14]))
                else:
                    self.xl_s6.append(None)

                if rows[15] is not None and rows[15] != '':
                    self.xl_s7.append(int(rows[15]))
                else:
                    self.xl_s7.append(None)

                if rows[16] is not None and rows[16] != '':
                    self.xl_s8.append(int(rows[16]))
                else:
                    self.xl_s8.append(None)

                if rows[17] is not None and rows[17] != '':
                    self.xl_s9.append(int(rows[17]))
                else:
                    self.xl_s9.append(None)
        except IOError:
            print("File not found or path is incorrect")

    def group_excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Group_Input_sheet'])
            sheet3 = workbook.sheet_by_index(0)
            for i in range(1, sheet3.nrows):
                number = i  # Counting number of rows
                rows = sheet3.row_values(number)

                # self.xl_GsCandidate.append(int(rows[0]))

                if rows[9] is not None and rows[9] != '':
                    self.xl_g1.append(int(rows[9]))
                else:
                    self.xl_g1.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_g2.append(int(rows[10]))
                else:
                    self.xl_g2.append(None)

                if rows[11] is not None and rows[11] != '':
                    self.xl_g3.append(int(rows[11]))
                else:
                    self.xl_g3.append(None)

                if rows[12] is not None and rows[12] != '':
                    self.xl_g4.append(int(rows[12]))
                else:
                    self.xl_g4.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_TotalMarks.append(int(rows[13]))
                else:
                    self.xl_TotalMarks.append(None)

        except IOError:
            print("File not found or path is incorrect")

    def upload_sheet(self, loop):

        self.lambda_function('uploadCandidatesScore')
        self.headers['APP-NAME'] = 'crpo'

        # ------------------          ---------------------------------------------
        # Upload Score Sheet ******** Every 30 Days S3 "FilePath" has to be Replace
        # ------------------          ---------------------------------------------
        uploadsheetrequest = {
            "TestId": self.xl_testId[loop],
            "FilePath": "https://s3-ap-southeast-1.amazonaws.com/test-all-hirepro-files/Automation/"
                        "assessmentScoreSheets/9d07816d-5d7e-4d07-8205-314f107e3c0fGroup_Section.xlsx",
            "Sync": "False"
        }
        uploadsheet_api = requests.post(self.webapi, headers=self.headers,
                                        data=json.dumps(uploadsheetrequest, default=str), verify=False)
        print(uploadsheet_api.headers)
        print(uploadsheet_api.headers)
        upload_api_dict = json.loads(uploadsheet_api.content)
        print(upload_api_dict)

    def updated_group_section_excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Update_group_section_input_sheet'])
            sheet4 = workbook.sheet_by_index(0)
            for i in range(1, sheet4.nrows):
                number = i  # Counting number of rows
                rows = sheet4.row_values(number)

                # self.xl_GsCandidate.append(int(rows[0]))

                if rows[9] is not None and rows[9] != '':
                    self.xl_s1_updated.append(int(rows[9]))
                else:
                    self.xl_s1_updated.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_s2_updated.append(int(rows[10]))
                else:
                    self.xl_s2_updated.append(None)

                if rows[11] is not None and rows[11] != '':
                    self.xl_s3_updated.append(int(rows[11]))
                else:
                    self.xl_s3_updated.append(None)

                if rows[12] is not None and rows[12] != '':
                    self.xl_s4_updated.append(int(rows[12]))
                else:
                    self.xl_s4_updated.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_s5_updated.append(int(rows[13]))
                else:
                    self.xl_s5_updated.append(None)

                if rows[14] is not None and rows[14] != '':
                    self.xl_s6_updated.append(int(rows[14]))
                else:
                    self.xl_s6_updated.append(None)

                if rows[15] is not None and rows[15] != '':
                    self.xl_s7_updated.append(int(rows[15]))
                else:
                    self.xl_s7_updated.append(None)

                if rows[16] is not None and rows[16] != '':
                    self.xl_s8_updated.append(int(rows[16]))
                else:
                    self.xl_s8_updated.append(None)

                if rows[17] is not None and rows[17] != '':
                    self.xl_s9_updated.append(int(rows[17]))
                else:
                    self.xl_s9_updated.append(None)
        except IOError:
            print("File not found or path is incorrect")

    def updated_group_excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Updated_group_Input_sheet'])
            sheet5 = workbook.sheet_by_index(0)
            for i in range(1, sheet5.nrows):
                number = i  # Counting number of rows
                rows = sheet5.row_values(number)

                # self.xl_GsCandidate.append(int(rows[0]))

                if not rows[9]:
                    self.xl_g1_updated.append(None)
                else:
                    self.xl_g1_updated.append(int(rows[9]))

                if not rows[10]:
                    self.xl_g2_updated.append(None)
                else:
                    self.xl_g2_updated.append(int(rows[10]))

                if not rows[11]:
                    self.xl_g3_updated.append(None)
                else:
                    self.xl_g3_updated.append(int(rows[11]))

                if not rows[12]:
                    self.xl_g4_updated.append(None)
                else:
                    self.xl_g4_updated.append(int(rows[12]))

                if not rows[13]:
                    self.xl_TotalMarks_updated.append(None)
                else:
                    self.xl_TotalMarks_updated.append(int(rows[13]))

        except IOError:
            print("File not found or path is incorrect")

    def updated_upload_sheet(self, loop):

        self.lambda_function('uploadCandidatesScore')
        self.headers['APP-NAME'] = 'crpo'

        # ------------------          ---------------------------------------
        # Upload Score Sheet ******** Every 30 Days Replace the S3 "FilePath"
        # ------------------          ---------------------------------------
        uploadsheetrequest = {
            "TestId": self.xl_testId[loop],
            "FilePath": "https://s3-ap-southeast-1.amazonaws.com/test-all-hirepro-files/Automation/"
                        "assessmentScoreSheets/98b269f2-e8fb-4a6c-98e2-7feb224de48cUpdated_Group_Section.xlsx",
            "Sync": "False"
        }
        uploadsheet_api = requests.post(self.webapi, headers=self.headers,
                                        data=json.dumps(uploadsheetrequest, default=str), verify=False)
        print(uploadsheet_api.headers)
        upload_api_dict = json.loads(uploadsheet_api.content)
        print(upload_api_dict)

    def fetching_scores(self, loop):

        self.lambda_function('getApplicantsInfo')
        self.headers['APP-NAME'] = 'crpo'

        score_request = {
            "CandidateIds": [self.xl_candidateId[loop]]
        }
        fetchingscores_api = requests.post(self.webapi, headers=self.headers,
                                           data=json.dumps(score_request, default=str), verify=False)
        print(fetchingscores_api.headers)
        time.sleep(5.2)
        fetchingscores_dict = json.loads(fetchingscores_api.content)
        scoredata = fetchingscores_dict.get('data')
        for testuser in scoredata:
            if testuser['CandidateId'] == self.xl_candidateId[loop]:
                self.testuser_dict = next(
                    (item for item in scoredata if item['CandidateId'] == self.xl_candidateId[loop]), None)
                # print testuser_dict

                assessment_dict = self.testuser_dict['AssessmentDetails']
                print(assessment_dict)
                for permission_to_go in assessment_dict:
                    self.is_offline = permission_to_go['IsOffline']
                    if permission_to_go['TestStatus'] == "NotAttended":
                        self.candidate_info_dict = next(
                            (item for item in assessment_dict if item['Id'] == self.xl_testId[loop]), None)
                    else:

                        for test_user in assessment_dict:
                            if test_user['Id'] == self.xl_testId[loop]:
                                self.test_details_dict = next(
                                    (item for item in assessment_dict if item['Id'] == self.xl_testId[loop]), None)
                                print(self.test_details_dict)
                                for group_details in self.test_details_dict['GroupWiseInfo']:

                                    # ------------
                                    # Group - 1
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group1[loop]:
                                        self.group1_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group1[loop]), None)
                                        self.section1_dict = next(
                                            (item for item in
                                             self.group1_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section1[loop]), None)
                                        self.section1_1_dict = next(
                                            (item for item in
                                             self.group1_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section1_1[loop]), None)
                                        print(self.section1_1_dict)
                                        self.section1_2_dict = next(
                                            (item for item in
                                             self.group1_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section1_2[loop]), None)
                                        # print self.section1_dict
                                        # print self.section1_1_dict
                                        # print self.section1_2_dict
                                        # print self.group1_dict

                                    # ------------
                                    # Group - 2
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group2[loop]:
                                        self.group2_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group2[loop]), None)
                                        self.section2_dict = next(
                                            (item for item in
                                             self.group2_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section2[loop]), None)
                                        self.section2_1_dict = next(
                                            (item for item in
                                             self.group2_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section2_1[loop]), None)
                                        # print self.section2_dict
                                        # print self.section2_1_dict
                                        # print self.group2_dict

                                    # ------------
                                    # Group - 3
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group3[loop]:
                                        self.group3_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group3[loop]), None)
                                        self.section3_dict = next(
                                            (item for item in
                                             self.group3_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section3[loop]), None)
                                        self.section3_1_dict = next(
                                            (item for item in
                                             self.group3_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section3_1[loop]), None)
                                        # print self.section3_dict
                                        # print self.section3_1_dict
                                        # print self.group3_dict

                                    # ------------
                                    # Group - 4
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group4[loop]:
                                        self.group4_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group4[loop]), None)
                                        self.section4_dict = next(
                                            (item for item in
                                             self.group4_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section4[loop]), None)
                                        self.section4_1_dict = next(
                                            (item for item in
                                             self.group4_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section4_1[loop]), None)
                                        # print self.section3_dict
                                        # print self.section3_1_dict
                                        # print self.group3_dict

    def output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_TotalMarks[loop] is not None and self.xl_TotalMarks[loop] >= 0:
            self.ws.write(self.rowsize, 7, self.xl_TotalMarks[loop])
        else:
            self.ws.write(self.rowsize, 7, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g1[loop] is not None and self.xl_g1[loop] >= 0:
            self.ws.write(self.rowsize, 8, self.xl_g1[loop])
        else:
            self.ws.write(self.rowsize, 8, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s1[loop] is not None and self.xl_s1[loop] >= 0:
            self.ws.write(self.rowsize, 9, self.xl_s1[loop])
        else:
            self.ws.write(self.rowsize, 9, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s2[loop] is not None and self.xl_s2[loop] >= 0:
            self.ws.write(self.rowsize, 10, self.xl_s2[loop])
        else:
            self.ws.write(self.rowsize, 10, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s3[loop]is not None and self.xl_s3[loop] >= 0:
            self.ws.write(self.rowsize, 11, self.xl_s3[loop])
        else:
            self.ws.write(self.rowsize, 11, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g2[loop] is not None and self.xl_g2[loop] >= 0:
            self.ws.write(self.rowsize, 12, self.xl_g2[loop])
        else:
            self.ws.write(self.rowsize, 12, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s4[loop] is not None and self.xl_s4[loop] >= 0:
            self.ws.write(self.rowsize, 13, self.xl_s4[loop])
        else:
            self.ws.write(self.rowsize, 13, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s5[loop] is not None and self.xl_s5[loop] >= 0:
            self.ws.write(self.rowsize, 14, self.xl_s5[loop])
        else:
            self.ws.write(self.rowsize, 14, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g3[loop] is not None and self.xl_g3[loop] >= 0:
            self.ws.write(self.rowsize, 15, self.xl_g3[loop])
        else:
            self.ws.write(self.rowsize, 15, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s6[loop] is not None and self.xl_s6[loop] >= 0:
            self.ws.write(self.rowsize, 16, self.xl_s6[loop])
        else:
            self.ws.write(self.rowsize, 16, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s7[loop] is not None and self.xl_s7[loop] >= 0:
            self.ws.write(self.rowsize, 17, self.xl_s7[loop])
        else:
            self.ws.write(self.rowsize, 17, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g4[loop] is not None and self.xl_g4[loop] >= 0:
            self.ws.write(self.rowsize, 18, self.xl_g4[loop])
        else:
            self.ws.write(self.rowsize, 18, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s8[loop] is not None and self.xl_s8[loop] >= 0:
            self.ws.write(self.rowsize, 19, self.xl_s8[loop])
        else:
            self.ws.write(self.rowsize, 19, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s9[loop] is not None and self.xl_s9[loop] >= 0:
            self.ws.write(self.rowsize, 20, self.xl_s9[loop])
        else:
            self.ws.write(self.rowsize, 20, 'Empty')

        # -------------------
        # Writing Output data
        # -------------------

        self.Positive_Status = 'Pass'

        if self.xl_s1[loop] != self.section1_dict.get('CandidateScoreTotal') if self.section1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s2[loop] != self.section1_1_dict.get('CandidateScoreTotal') if self.section1_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s3[loop] != self.section1_2_dict.get('CandidateScoreTotal') if self.section1_2_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s4[loop] != self.section2_dict.get('CandidateScoreTotal') if self.section2_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s5[loop] != self.section2_1_dict.get('CandidateScoreTotal') if self.section2_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s6[loop] != self.section3_dict.get('CandidateScoreTotal') if self.section3_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s7[loop] != self.section3_1_dict.get('CandidateScoreTotal') if self.section3_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s8[loop] != self.section4_dict.get('CandidateScoreTotal') if self.section4_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s9[loop] != self.section4_1_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g1[loop] != self.group1_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g2[loop] != self.group2_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g3[loop] != self.group3_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g4[loop] != self.group4_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'

        if self.Positive_Status == 'Fail':
            self.Negitive_Status = 'Fail'

        self.rowsize += 1  # Row increment

        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        if self.testuser_dict and self.testuser_dict.get('CandidateId'):
            self.ws.write(self.rowsize, 2, self.testuser_dict.get('CandidateId', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.testuser_dict and self.testuser_dict.get('CandidateName'):
            self.ws.write(self.rowsize, 3, self.testuser_dict.get('CandidateName', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.testuser_dict and self.testuser_dict.get('Email'):
            self.ws.write(self.rowsize, 4, self.testuser_dict.get('Email', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.Positive_Status == 'Pass':
            self.ws.write(self.rowsize, 1, self.Positive_Status, self.style8)
        else:
            self.ws.write(self.rowsize, 1, self.Positive_Status, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_offline:
            self.ws.write(self.rowsize, 5, "Offline")
        else:
            self.ws.write(self.rowsize, 5, "Online")
        # --------------------------------------------------------------------------------------------------------------

        if self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 6, self.candidate_info_dict.get('TestStatus'))
        else:
            self.ws.write(self.rowsize, 6, 'Attended')
        # --------------------------------------------------------------------------------------------------------------

        if self.test_details_dict and self.test_details_dict.get('CandidateMarks'):
            self.ws.write(self.rowsize, 7, self.test_details_dict.get('CandidateMarks'), self.style14)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 7, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 7, self.test_details_dict.get('CandidateMarks'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.group1_dict and self.group1_dict.get('CandidateScoreTotal'):
            if self.xl_g1[loop] == self.group1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 8, self.group1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 8, self.group1_dict.get('CandidateScoreTotal'), self.style3)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 8, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_g1[loop] is None:
            self.ws.write(self.rowsize, 8, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_dict and self.section1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s1[loop] == self.section1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 9, self.section1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 9, self.section1_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 9, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s1[loop] is None:
            self.ws.write(self.rowsize, 9, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_1_dict and self.section1_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s2[loop] == self.section1_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 10, self.section1_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 10, self.section1_1_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 10, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s2[loop] is None:
            self.ws.write(self.rowsize, 10, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_2_dict and self.section1_2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s3[loop] == self.section1_2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 11, self.section1_2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 11, self.section1_2_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 11, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s3[loop] is None:
            self.ws.write(self.rowsize, 11, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.group2_dict and self.group2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g2[loop] == self.group2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 12, self.group2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 12, self.group2_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 12, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_g2[loop] is None:
            self.ws.write(self.rowsize, 12, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section2_dict and self.section2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s4[loop] == self.section2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 13, self.section2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 13, self.section2_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 13, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s4[loop] is None:
            self.ws.write(self.rowsize, 13, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section2_1_dict and self.section2_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s5[loop] == self.section2_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 14, self.section2_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 14, self.section2_1_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 14, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s5[loop] is None:
            self.ws.write(self.rowsize, 14, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.group3_dict and self.group3_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g3[loop] == self.group3_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 15, self.group3_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 15, self.group3_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 15, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_g3[loop] is None:
            self.ws.write(self.rowsize, 15, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section3_dict and self.section3_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s6[loop] == self.section3_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 16, self.section3_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 16, self.section3_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 16, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s6[loop] is None:
            self.ws.write(self.rowsize, 16, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section3_1_dict and self.section3_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s7[loop] == self.section3_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 17, self.section3_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 17, self.section3_1_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 17, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s7[loop] is None:
            self.ws.write(self.rowsize, 17, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.group4_dict and self.group4_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g4[loop] == self.group4_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 18, self.group4_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 18, self.group4_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 18, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_g4[loop] is None:
            self.ws.write(self.rowsize, 18, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section4_dict and self.section4_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s8[loop] == self.section4_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 19, self.section4_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 19, self.section4_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 19, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s8[loop] is None:
            self.ws.write(self.rowsize, 19, "Empty", self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.section4_1_dict and self.section4_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s9[loop] == self.section4_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize, 20, self.section4_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize, 20, self.section4_1_dict.get('CandidateScoreTotal'), self.style18)
        elif self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize, 20, 'Empty', self.style14)
        elif self.candidate_info_dict.get('TestStatus') != 'NotAttended' and self.xl_s9[loop] is None:
            self.ws.write(self.rowsize, 20, "Empty", self.style14)

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Score_Output_sheet'])

    def updated_output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        self.rowsize1 += 1
        self.ws.write(self.rowsize1, self.col, 'Input', self.style4)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_TotalMarks_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 7, self.xl_TotalMarks_updated[loop])
        else:
            self.ws.write(self.rowsize1, 7, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g1_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 8, self.xl_g1_updated[loop])
        else:
            self.ws.write(self.rowsize1, 8, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s1_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 9, self.xl_s1_updated[loop])
        else:
            self.ws.write(self.rowsize1, 9, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s2_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 10, self.xl_s2_updated[loop])
        else:
            self.ws.write(self.rowsize1, 10, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s3_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 11, self.xl_s3_updated[loop])
        else:
            self.ws.write(self.rowsize1, 11, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g2_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 12, self.xl_g2_updated[loop])
        else:
            self.ws.write(self.rowsize1, 12, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s4_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 13, self.xl_s4_updated[loop])
        else:
            self.ws.write(self.rowsize1, 13, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s5_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 14, self.xl_s5_updated[loop])
        else:
            self.ws.write(self.rowsize1, 14, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g3_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 15, self.xl_g3_updated[loop])
        else:
            self.ws.write(self.rowsize1, 15, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s6_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 16, self.xl_s6_updated[loop])
        else:
            self.ws.write(self.rowsize1, 16, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s7_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 17, self.xl_s7_updated[loop])
        else:
            self.ws.write(self.rowsize1, 17, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_g4_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 18, self.xl_g4_updated[loop])
        else:
            self.ws.write(self.rowsize1, 18, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s8_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 19, self.xl_s8_updated[loop])
        else:
            self.ws.write(self.rowsize1, 19, 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_s9_updated[loop] >= 0:
            self.ws.write(self.rowsize1, 20, self.xl_s9_updated[loop])
        else:
            self.ws.write(self.rowsize1, 20, 'Empty')

        # -------------------
        # Writing Output data
        # -------------------
        # self.rowsize += 1  # Row increment
        self.Positive_Status = 'Pass'

        if self.xl_s1_updated[loop] != self.section1_dict.get('CandidateScoreTotal') if self.section1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s2_updated[loop] != self.section1_1_dict.get('CandidateScoreTotal') if self.section1_1_dict \
                else None:
            self.Positive_Status = 'Fail'
        if self.xl_s3_updated[loop] != self.section1_2_dict.get('CandidateScoreTotal') if self.section1_2_dict \
                else None:
            self.Positive_Status = 'Fail'
        if self.xl_s4_updated[loop] != self.section2_dict.get('CandidateScoreTotal') if self.section2_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s5_updated[loop] != self.section2_1_dict.get('CandidateScoreTotal') if self.section2_1_dict \
                else None:
            self.Positive_Status = 'Fail'
        if self.xl_s6_updated[loop] != self.section3_dict.get('CandidateScoreTotal') if self.section3_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s7_updated[loop] != self.section3_1_dict.get('CandidateScoreTotal') if self.section3_1_dict \
                else None:
            self.Positive_Status = 'Fail'
        if self.xl_s8_updated[loop] != self.section4_dict.get('CandidateScoreTotal') if self.section4_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_s9_updated[loop] != self.section4_1_dict.get('CandidateScoreTotal') if self.section4_1_dict \
                else None:
            self.Positive_Status = 'Fail'
        if self.xl_g1_updated[loop] != self.group1_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g2_updated[loop] != self.group2_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g3_updated[loop] != self.group3_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'
        if self.xl_g4_updated[loop] != self.group4_dict.get('CandidateScoreTotal') if self.section4_1_dict else None:
            self.Positive_Status = 'Fail'

        if self.Positive_Status == 'Fail':
            self.Negitive_Status = 'Fail'

        self.ws.write(self.rowsize1 + 1, self.col, 'Output', self.style5)

        if self.testuser_dict and self.testuser_dict.get('CandidateId'):
            self.ws.write(self.rowsize1 + 1, 2, self.testuser_dict.get('CandidateId', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.testuser_dict and self.testuser_dict.get('CandidateName'):
            self.ws.write(self.rowsize1 + 1, 3, self.testuser_dict.get('CandidateName', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.testuser_dict and self.testuser_dict.get('Email'):
            self.ws.write(self.rowsize1 + 1, 4, self.testuser_dict.get('Email', None))
        # --------------------------------------------------------------------------------------------------------------

        if self.Positive_Status == 'Pass':
            self.ws.write(self.rowsize1 + 1, 1, self.Positive_Status, self.style8)
        else:
            self.ws.write(self.rowsize1 + 1, 1, self.Positive_Status, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_offline:
            self.ws.write(self.rowsize1 + 1, 5, "Offline")
        else:
            self.ws.write(self.rowsize1 + 1, 5, "Online")
        # --------------------------------------------------------------------------------------------------------------
        if self.candidate_info_dict.get('TestStatus') == 'NotAttended':
            self.ws.write(self.rowsize1 + 1, 6, self.candidate_info_dict.get('TestStatus'))
        else:
            self.ws.write(self.rowsize1 + 1, 6, 'Attended')
        # --------------------------------------------------------------------------------------------------------------

        if self.test_details_dict and self.test_details_dict.get('CandidateMarks'):
            self.ws.write(self.rowsize1 + 1, 7, self.test_details_dict.get('CandidateMarks'), self.style14)
        else:
            self.ws.write(self.rowsize1 + 1, 7, self.test_details_dict.get('CandidateMarks'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.group1_dict and self.group1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g1_updated[loop] == self.group1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 8, self.group1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 8, self.group1_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_dict and self.section1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s1_updated[loop] == self.section1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 9, self.section1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 9, self.section1_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_1_dict and self.section1_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s2_updated[loop] == self.section1_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 10, self.section1_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 10, self.section1_1_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section1_2_dict and self.section1_2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s3_updated[loop] == self.section1_2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 11,
                              self.section1_2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 11,
                              self.section1_2_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.group2_dict and self.group2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g2_updated[loop] == self.group2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 12, self.group2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 12, self.group2_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section2_dict and self.section2_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s4_updated[loop] == self.section2_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 13, self.section2_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 13, self.section2_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section2_1_dict and self.section2_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s5_updated[loop] == self.section2_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 14,
                              self.section2_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 14,
                              self.section2_1_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.group3_dict and self.group3_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g3_updated[loop] == self.group3_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 15, self.group3_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 15, self.group3_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section3_dict and self.section3_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s6_updated[loop] == self.section3_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 16, self.section3_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 16, self.section3_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section3_1_dict and self.section3_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s7_updated[loop] == self.section3_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 17,
                              self.section3_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 17,
                              self.section3_1_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.group4_dict and self.group4_dict.get('CandidateScoreTotal') is not None:
            if self.xl_g4_updated[loop] == self.group4_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 18, self.group4_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 18, self.group4_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section4_dict and self.section4_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s8_updated[loop] == self.section4_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 19, self.section4_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 19, self.section4_dict.get('CandidateScoreTotal'), self.style18)
        # --------------------------------------------------------------------------------------------------------------

        if self.section4_1_dict and self.section4_1_dict.get('CandidateScoreTotal') is not None:
            if self.xl_s9_updated[loop] == self.section4_1_dict.get('CandidateScoreTotal'):
                self.ws.write(self.rowsize1 + 1, 20,
                              self.section4_1_dict.get('CandidateScoreTotal'), self.style14)
            else:
                self.ws.write(self.rowsize1 + 1, 20,
                              self.section4_1_dict.get('CandidateScoreTotal'), self.style18)

        self.rowsize1 += 1
        Object.wb_Result.save(output_paths.outputpaths['Score_Output_sheet'])

    def margin_line(self):
        self.ws.write(16, self.col, '.', self.style9)
        self.ws.write(16, 1, '.', self.style9)
        self.ws.write(16, 2, '.', self.style9)
        self.ws.write(16, 3, '.', self.style9)
        self.ws.write(16, 4, '.', self.style9)
        self.ws.write(16, 5, '.', self.style9)
        self.ws.write(16, 6, '.', self.style9)
        self.ws.write(16, 7, '.', self.style9)
        self.ws.write(16, 8, '.', self.style9)
        self.ws.write(16, 9, '.', self.style9)
        self.ws.write(16, 10, '.', self.style9)
        self.ws.write(16, 11, '.', self.style9)
        self.ws.write(16, 12, '.', self.style9)
        self.ws.write(16, 13, '.', self.style9)
        self.ws.write(16, 14, '.', self.style9)
        self.ws.write(16, 15, '.', self.style9)
        self.ws.write(16, 16, '.', self.style9)
        self.ws.write(16, 17, '.', self.style9)
        self.ws.write(16, 18, '.', self.style9)
        self.ws.write(16, 19, '.', self.style9)
        self.ws.write(16, 20, '.', self.style9)
        self.ws.write(17, self.col, '.', self.style9)
        self.ws.write(17, 1, '.', self.style9)
        self.ws.write(17, 2, '.', self.style9)
        self.ws.write(17, 3, '.', self.style9)
        self.ws.write(17, 4, '.', self.style9)
        self.ws.write(17, 5, 'UPDATED', self.style9)
        self.ws.write(17, 6, '.', self.style9)
        self.ws.write(17, 7, '.', self.style9)
        self.ws.write(17, 8, '.', self.style9)
        self.ws.write(17, 9, '.', self.style9)
        self.ws.write(17, 10, '.', self.style9)
        self.ws.write(17, 11, '.', self.style9)
        self.ws.write(17, 12, '.', self.style9)
        self.ws.write(17, 13, '.', self.style9)
        self.ws.write(17, 14, '.', self.style9)
        self.ws.write(17, 15, '.', self.style9)
        self.ws.write(17, 16, '.', self.style9)
        self.ws.write(17, 17, '.', self.style9)
        self.ws.write(17, 18, '.', self.style9)
        self.ws.write(17, 19, '.', self.style9)
        self.ws.write(17, 20, '.', self.style9)
        self.ws.write(18, self.col, '.', self.style9)
        self.ws.write(18, 1, '.', self.style9)
        self.ws.write(18, 2, '.', self.style9)
        self.ws.write(18, 3, '.', self.style9)
        self.ws.write(18, 4, '.', self.style9)
        self.ws.write(18, 5, '.', self.style9)
        self.ws.write(18, 6, '.', self.style9)
        self.ws.write(18, 7, '.', self.style9)
        self.ws.write(18, 8, '.', self.style9)
        self.ws.write(18, 9, '.', self.style9)
        self.ws.write(18, 10, '.', self.style9)
        self.ws.write(18, 11, '.', self.style9)
        self.ws.write(18, 12, '.', self.style9)
        self.ws.write(18, 13, '.', self.style9)
        self.ws.write(18, 14, '.', self.style9)
        self.ws.write(18, 15, '.', self.style9)
        self.ws.write(18, 16, '.', self.style9)
        self.ws.write(18, 17, '.', self.style9)
        self.ws.write(18, 18, '.', self.style9)
        self.ws.write(18, 19, '.', self.style9)
        self.ws.write(18, 20, '.', self.style9)

    def overall_status(self):
        self.ws.write(self.final_status_rowsize, 0, 'Upload Score', self.style23)
        if self.Positive_Status == 'Pass':
            self.ws.write(self.final_status_rowsize, 1, 'Pass', self.style24)
        else:
            self.ws.write(self.final_status_rowsize, 1, self.Negitive_Status, self.style25)
        self.ws.write(self.final_status_rowsize, 2, 'StartTime', self.style23)
        self.ws.write(self.final_status_rowsize, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, NumberOfTestCases, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Score_Output_sheet'])


Object = UploadScoresheet()
Object.excel_headers()
Object.excel_data()
Object.group_section_excel_data()
Object.updated_group_section_excel_data()
Object.group_excel_data()
Object.updated_group_excel_data()
Object.margin_line()

Total_count = len(Object.xl_candidateId)
print("Number Of Rows ::", Total_count)
Total_count1 = len(Object.xl_s1_updated)
print("Number Of Rows ::", Total_count1)
NumberOfTestCases = Total_count + Total_count1

if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.upload_sheet(looping)
        Object.fetching_scores(looping)
        Object.output_excel(looping)

        # -------------------------------------
        # Making all dict empty for every loop
        # -------------------------------------
        Object.candidate_info_dict = {}
        Object.testuser_dict = {}
        Object.test_details_dict = {}
        Object.group1_dict = {}
        Object.group2_dict = {}
        Object.group3_dict = {}
        Object.group4_dict = {}
        Object.section1_dict = {}
        Object.section2_dict = {}
        Object.section3_dict = {}
        Object.section4_dict = {}
        Object.section1_1_dict = {}
        Object.section1_2_dict = {}
        Object.section2_1_dict = {}
        Object.section3_1_dict = {}
        Object.section4_1_dict = {}

    for looping in range(0, Total_count1):
        print("Iteration Count is ::", looping)
        Object.updated_upload_sheet(looping)
        Object.fetching_scores(looping)
        Object.updated_output_excel(looping)

        # -------------------------------------
        # Making all dict empty for every loop
        # -------------------------------------
        Object.candidate_info_dict = {}
        Object.testuser_dict = {}
        Object.test_details_dict = {}
        Object.group1_dict = {}
        Object.group2_dict = {}
        Object.group3_dict = {}
        Object.group4_dict = {}
        Object.section1_dict = {}
        Object.section2_dict = {}
        Object.section3_dict = {}
        Object.section4_dict = {}
        Object.section1_1_dict = {}
        Object.section1_2_dict = {}
        Object.section2_1_dict = {}
        Object.section3_1_dict = {}
        Object.section4_1_dict = {}
Object.overall_status()
