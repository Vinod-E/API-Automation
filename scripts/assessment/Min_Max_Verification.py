from hpro_automation import (login, work_book, input_paths, output_paths, db_login)
import datetime
import requests
import json
import xlrd


class Min_Max_Verification(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):
        # ---------------------------------- Overall Status Run Date ---------------------------------------------------
        self.start_time = str(datetime.datetime.now())

        # --------------------------------- Inheritance Class Instance -------------------------------------------------
        super(Min_Max_Verification, self).__init__()
        self.common_login('automation')
        self.db_connection('amsin')

        # --------------------------------- Overall status initialize variables ----------------------------------------
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 4)))
        self.Actual_Success_case = []
        self.Expected_success_cases_each = list(map(lambda x: 'Pass', range(0, 29)))
        self.Actual_Success_case_each = []

        # --------------------------------- Excel Data initialize variables --------------------------------------------
        self.xl_testId = []
        self.xl_G1_Name = []
        self.xl_G1_Max = []
        self.xl_G1_Min = []
        self.xl_G1_S1_Name = []
        self.xl_G1_S1_Max = []
        self.xl_G1_S1_Min = []
        self.xl_G1_S2_Name = []
        self.xl_G1_S2_Max = []
        self.xl_G1_S2_Min = []
        self.xl_G2_Name = []
        self.xl_G2_Max = []
        self.xl_G2_Min = []
        self.xl_G2_S1_Name = []
        self.xl_G2_S1_Max = []
        self.xl_G2_S1_Min = []
        self.xl_G2_S2_Name = []
        self.xl_G2_S2_Max = []
        self.xl_G2_S2_Min = []
        self.xl_G3_Name = []
        self.xl_G3_Max = []
        self.xl_G3_Min = []
        self.xl_G3_S1_Name = []
        self.xl_G3_S1_Max = []
        self.xl_G3_S1_Min = []
        self.xl_G3_S2_Name = []
        self.xl_G3_S2_Max = []
        self.xl_G3_S2_Min = []
        self.xl_Test_Max = []
        self.xl_Test_Min = []

        # --------------------------------- Dictionary initialize variables --------------------------------------------
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.headers = {}

    def excel_headers(self):
        # --------------------------------- Excel Headers and Cell color, styles ---------------------------------------
        self.main_headers = ['Comparison', 'status', 'Test Id',
                             'G1_Name','G1_Max','G1_Min',
                             'G1_S1_Name','G1_S1_Max','G1_S1_Min',
                             'G1_s2_Name','G1_S2_Max','G1_S2_Min',
                             'G2_Name','G2_Max','G2_Min',
                             'G2_S1_Name','G2_S1_Max','G2_S1_Min',
                             'G2_S2_Name','G2_S2_Max','G2_S2_Min',
                             'G3_Name','G3_Max','G3_Min',
                             'G3_S1_Name','G3_S1_Max','G3_S1_Min',
                             'G3_S2_Name','G3_S2_Max','G3_S2_Min',
                             'Test_Min','Test_Max']
        self.headers_with_style2 = ['Test Id','G1_Name','G1_Max','G1_Min',
                             'G1_S1_Name','G1_S1_Max','G1_S1_Min',
                             'G1_s2_Name','G1_S2_Max','G1_S2_Min',
                             'G2_Name','G2_Max','G2_Min',
                             'G2_S1_Name','G2_S1_Max','G2_S1_Min',
                             'G2_S2_Name','G2_S2_Max','G2_S2_Min',
                             'G3_Name','G3_Max','G3_Min',
                             'G3_S1_Name','G3_S1_Max','G3_S1_Min',
                             'G3_S2_Name','G3_S2_Max','G3_S2_Min',
                             'Test_Min','Test_Max']
        self.file_headers_col_row()

    def excel_data(self):
        # ------------------------------- Excel Data Read --------------------------------------------------------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['min_max_Input_sheet'])
            sheet1 = workbook.sheet_by_index(0)
            for i in range(1, sheet1.nrows):
                number = i  # Counting number of rows
                rows = sheet1.row_values(number)
                self.xl_testId.append(rows[0])
                if not rows[1]:
                    self.xl_G1_Name.append(None)
                else:
                    self.xl_G1_Name.append(rows[1])
                if not rows[2]:
                    self.xl_G1_Max.append(None)
                else:
                    self.xl_G1_Max.append(rows[2])
                if not rows[3]:
                    self.xl_G1_Min.append(None)
                else:
                    self.xl_G1_Min.append(rows[3])
                if not rows[4]:
                    self.xl_G1_S1_Name.append(None)
                else:
                    self.xl_G1_S1_Name.append(rows[4])
                if not rows[5]:
                    self.xl_G1_S1_Max.append(None)
                else:
                    self.xl_G1_S1_Max.append(rows[5])
                if not rows[6]:
                    self.xl_G1_S1_Min.append(None)
                else:
                    self.xl_G1_S1_Min.append(rows[6])
                if not rows[7]:
                    self.xl_G1_S2_Name.append(None)
                else:
                    self.xl_G1_S2_Name.append(rows[7])
                if not rows[8]:
                    self.xl_G1_S2_Max.append(None)
                else:
                    self.xl_G1_S2_Max.append(rows[8])
                if not rows[9]:
                    self.xl_G1_S2_Min.append(None)
                else:
                    self.xl_G1_S2_Min.append(rows[9])
                if not rows[10]:
                    self.xl_G2_Name.append(None)
                else:
                    self.xl_G2_Name.append(rows[10])
                if not rows[11]:
                    self.xl_G2_Max.append(None)
                else:
                    self.xl_G2_Max.append(rows[11])
                if not rows[12]:
                    self.xl_G2_Min.append(None)
                else:
                    self.xl_G2_Min.append(rows[12])
                if not rows[13]:
                    self.xl_G2_S1_Name.append(None)
                else:
                    self.xl_G2_S1_Name.append(rows[13])
                if not rows[14]:
                    self.xl_G2_S1_Max.append(None)
                else:
                    self.xl_G2_S1_Max.append(rows[14])
                if not rows[15]:
                    self.xl_G2_S1_Min.append(None)
                else:
                    self.xl_G2_S1_Min.append(rows[15])
                if not rows[16]:
                    self.xl_G2_S2_Name.append(None)
                else:
                    self.xl_G2_S2_Name.append(rows[16])
                if not rows[17]:
                    self.xl_G2_S2_Max.append(None)
                else:
                    self.xl_G2_S2_Max.append(rows[17])
                if not rows[18]:
                    self.xl_G2_S2_Min.append(None)
                else:
                    self.xl_G2_S2_Min.append(rows[18])
                if not rows[19]:
                    self.xl_G3_Name.append(None)
                else:
                    self.xl_G3_Name.append(rows[19])
                if not rows[20]:
                    self.xl_G3_Max.append(None)
                else:
                    self.xl_G3_Max.append(rows[20])
                if not rows[21]:
                    self.xl_G3_Min.append(None)
                else:
                    self.xl_G3_Min.append(rows[21])
                if not rows[22]:
                    self.xl_G3_S1_Name.append(None)
                else:
                    self.xl_G3_S1_Name.append(rows[22])
                if not rows[23]:
                    self.xl_G3_S1_Max.append(None)
                else:
                    self.xl_G3_S1_Max.append(rows[23])
                if not rows[24]:
                    self.xl_G3_S1_Min.append(None)
                else:
                    self.xl_G3_S1_Min.append(rows[24])
                if not rows[25]:
                    self.xl_G3_S2_Name.append(None)
                else:
                    self.xl_G3_S2_Name.append(rows[25])
                if not rows[26]:
                    self.xl_G3_S2_Max.append(None)
                else:
                    self.xl_G3_S2_Max.append(rows[26])
                if not rows[27]:
                    self.xl_G3_S2_Min.append(None)
                else:
                    self.xl_G3_S2_Min.append(rows[27])
                if not rows[28]:
                    self.xl_Test_Max.append(None)
                else:
                    self.xl_Test_Max.append(rows[28])
                if not rows[29]:
                    self.xl_Test_Min.append(None)
                else:
                    self.xl_Test_Min.append(rows[29])
        except IOError:
            print("File not found or path is incorrect")

    def api_call(self, loop):
        self.lambda_function('unApproveTest')
        self.headers['APP-NAME'] = 'crpo'
        unApprove_Request = {"testId": self.xl_testId[loop]}
        unApprove_hit_api = requests.post(self.webapi, headers=self.headers, data=json.dumps(unApprove_Request, default=str), verify=False)
        unApprove_api_response = json.loads(unApprove_hit_api.content)
        print('unApprove_api_response', unApprove_api_response)

        self.lambda_function('approveTest')
        self.headers['APP-NAME'] = 'crpo'
        approve_Request = {"testId": self.xl_testId[loop]}
        approve_hit_api = requests.post(self.webapi, headers=self.headers,
                                          data=json.dumps(approve_Request, default=str), verify=False)
        approve_api_response = json.loads(approve_hit_api.content)
        print('approve_api_response', approve_api_response)

        self.lambda_function('getTest')
        self.headers['APP-NAME'] = 'crpo'
        # ----------------------------------- API request --------------------------------------------------------------
        request = {"testId": self.xl_testId[loop]}
        hit_api = requests.post(self.webapi, headers=self.headers,
                                data=json.dumps(request, default=str), verify=False)
        hitted_api_response = json.loads(hit_api.content)
        data = hitted_api_response["data"]
        self.assessment_dict = data['assessment']
        group_dict = self.assessment_dict['testGroupSectionInfos']
        for grp_sec in group_dict:
            if grp_sec['groupName'] == self.xl_G1_Name[loop]:
                self.g1_dict = grp_sec
                for sec in self.g1_dict['testSectionInfo']:
                    if sec['sectionName'] == self.xl_G1_S1_Name[loop]:
                        self.g1_s1_dict = sec
                    elif sec['sectionName'] == self.xl_G1_S2_Name[loop]:
                        self.g1_s2_dict = sec
                    else:
                        print("No section found in group 1!!!!")
            elif grp_sec['groupName'] == self.xl_G2_Name[loop]:
                self.g2_dict = grp_sec
                for sec in self.g2_dict['testSectionInfo']:
                    if sec['sectionName'] == self.xl_G2_S1_Name[loop]:
                        self.g2_s1_dict = sec
                    elif sec['sectionName'] == self.xl_G2_S2_Name[loop]:
                        self.g2_s2_dict = sec
                    else:
                        print("No section found in group 2!!!!")
            elif grp_sec['groupName'] == self.xl_G3_Name[loop]:
                self.g3_dict = grp_sec
                for sec in self.g3_dict['testSectionInfo']:
                    if sec['sectionName'] == self.xl_G3_S1_Name[loop]:
                        self.g3_s1_dict = sec
                    elif sec['sectionName'] == self.xl_G3_S2_Name[loop]:
                        self.g3_s2_dict = sec
                    else:
                        print("No section found in group 3!!!!")
            else:
                print("No group found in test id", self.xl_testId[loop])

    def output_report(self, loop):
        # --------------------------------- Writing Input Data ---------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.xl_testId[loop] if self.xl_testId[loop] else 0)
        self.ws.write(self.rowsize, 3, self.xl_G1_Name[loop] if self.xl_G1_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 4, self.xl_G1_Max[loop] if self.xl_G1_Max[loop] else 0)
        self.ws.write(self.rowsize, 5, self.xl_G1_Min[loop] if self.xl_G1_Min[loop] else 0)
        self.ws.write(self.rowsize, 6, self.xl_G1_S1_Name[loop] if self.xl_G1_S1_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 7, self.xl_G1_S1_Max[loop] if self.xl_G1_S1_Max[loop] else 0)
        self.ws.write(self.rowsize, 8, self.xl_G1_S1_Min[loop] if self.xl_G1_S1_Min[loop] else 0)
        self.ws.write(self.rowsize, 9, self.xl_G1_S2_Name[loop] if self.xl_G1_S2_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 10, self.xl_G1_S2_Max[loop] if self.xl_G1_S2_Max[loop] else 0)
        self.ws.write(self.rowsize, 11, self.xl_G1_S2_Min[loop] if self.xl_G1_S2_Min[loop] else 0)
        self.ws.write(self.rowsize, 12, self.xl_G2_Name[loop] if self.xl_G2_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 13, self.xl_G2_Max[loop] if self.xl_G2_Max[loop] else 0)
        self.ws.write(self.rowsize, 14, self.xl_G2_Min[loop] if self.xl_G2_Min[loop] else 0)
        self.ws.write(self.rowsize, 15, self.xl_G2_S1_Name[loop] if self.xl_G2_S1_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 16, self.xl_G2_S1_Max[loop] if self.xl_G2_S1_Max[loop] else 0)
        self.ws.write(self.rowsize, 17, self.xl_G2_S1_Min[loop] if self.xl_G2_S1_Min[loop] else 0)
        self.ws.write(self.rowsize, 18, self.xl_G2_S2_Name[loop] if self.xl_G2_S2_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 19, self.xl_G2_S2_Max[loop] if self.xl_G2_S2_Max[loop] else 0)
        self.ws.write(self.rowsize, 20, self.xl_G2_S2_Min[loop] if self.xl_G2_S2_Min[loop] else 0)
        self.ws.write(self.rowsize, 21, self.xl_G3_Name[loop] if self.xl_G3_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 22, self.xl_G3_Max[loop] if self.xl_G3_Max[loop] else 0)
        self.ws.write(self.rowsize, 23, self.xl_G3_Min[loop] if self.xl_G3_Min[loop] else 0)
        self.ws.write(self.rowsize, 24, self.xl_G3_S1_Name[loop] if self.xl_G3_S1_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 25, self.xl_G3_S1_Max[loop] if self.xl_G3_S1_Max[loop] else 0)
        self.ws.write(self.rowsize, 26, self.xl_G3_S1_Min[loop] if self.xl_G3_S1_Min[loop] else 0)
        self.ws.write(self.rowsize, 27, self.xl_G3_S2_Name[loop] if self.xl_G3_S2_Name[loop] else 'Empty')
        self.ws.write(self.rowsize, 28, self.xl_G3_S2_Max[loop] if self.xl_G3_S2_Max[loop] else 0)
        self.ws.write(self.rowsize, 29, self.xl_G3_S2_Min[loop] if self.xl_G3_S2_Min[loop] else 0)
        self.ws.write(self.rowsize, 30, self.xl_Test_Min[loop] if self.xl_Test_Min[loop] else 0)
        self.ws.write(self.rowsize, 31, self.xl_Test_Max[loop] if self.xl_Test_Max[loop] else 0)
        self.rowsize += 1

        # --------------------------------- Writing Output Data --------------------------------------------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.assessment_dict['id'] == self.xl_testId[loop]:
            self.ws.write(self.rowsize, 2, int(self.assessment_dict['id']), self.style8)
        else:
            self.ws.write(self.rowsize, 2, int(self.assessment_dict['id']), self.style3)
        # ---------------------------------------------Group 1-------- -------------------------------------------------
        if self.g1_dict['groupName'] == self.xl_G1_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 3, self.g1_dict['groupName'], self.style8)
        else:
            self.ws.write(self.rowsize, 3, self.g1_dict['groupName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g1_dict['maxMark'] == self.xl_G1_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 4, self.g1_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 4, self.g1_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G1_Min[loop] is None:
            self.xl_G1_Min[loop] = 0
            if self.g1_dict['minMark'] == self.xl_G1_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 5, self.g1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 5, self.g1_dict['minMark'], self.style3)
        else:
            if round(-self.g1_dict['minMark'], 2) == round(self.xl_G1_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 5, -self.g1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 5, -self.g1_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g1_s1_dict['sectionName'] == self.xl_G1_S1_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 6, self.g1_s1_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 6, self.g1_s1_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g1_s1_dict['maxMark'] == self.xl_G1_S1_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 7, self.g1_s1_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 7, self.g1_s1_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G1_S1_Min[loop] is None:
            self.xl_G1_S1_Min[loop] = 0
            if self.g1_s1_dict['minMark'] == self.xl_G1_S1_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 8, self.g1_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 8, self.g1_s1_dict['minMark'], self.style3)
        else:
            if round(-self.g1_s1_dict['minMark'], 2) == round(self.xl_G1_S1_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 8, -self.g1_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 8, -self.g1_s1_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g1_s2_dict['sectionName'] == self.xl_G1_S2_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 9, self.g1_s2_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 9, self.g1_s2_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g1_s2_dict['maxMark'] == self.xl_G1_S2_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 10, self.g1_s2_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 10, self.g1_s2_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G1_S2_Min[loop] is None:
            self.xl_G1_S2_Min[loop] = 0
            if self.g1_s2_dict['minMark'] == self.xl_G1_S2_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 11, self.g1_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 11, self.g1_s2_dict['minMark'], self.style3)
        else:
            if round(-self.g1_s2_dict['minMark'], 2) == round(self.xl_G1_S2_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 11, -self.g1_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 11, -self.g1_s2_dict['minMark'], self.style3)
        # --------------------------------------------Group 2--------- -------------------------------------------------
        if self.g2_dict['groupName'] == self.xl_G2_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 12, self.g2_dict['groupName'], self.style8)
        else:
            self.ws.write(self.rowsize, 12, self.g2_dict['groupName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g2_dict['maxMark'] == self.xl_G2_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 13, self.g2_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 13, self.g2_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G2_Min[loop] is None:
            self.xl_G2_Min[loop] = 0
            if self.g2_dict['minMark'] == self.xl_G2_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 14, self.g2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 14, self.g2_dict['minMark'], self.style3)
        else:
            if round(-self.g2_dict['minMark'], 2) == round(self.xl_G2_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 14, -self.g2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 14, -self.g2_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g2_s1_dict['sectionName'] == self.xl_G2_S1_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 15, self.g2_s1_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 15, self.g2_s1_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g2_s1_dict['maxMark'] == self.xl_G2_S1_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 16, self.g2_s1_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 16, self.g2_s1_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G2_S1_Min[loop] is None:
            self.xl_G2_S1_Min[loop] = 0
            if self.g2_s1_dict['minMark'] == self.xl_G2_S1_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 17, self.g2_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 17, self.g2_s1_dict['minMark'], self.style3)
        else:
            if round(-self.g2_s1_dict['minMark'], 2) == round(self.xl_G2_S1_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 17, -self.g2_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 17, -self.g2_s1_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g2_s2_dict['sectionName'] == self.xl_G2_S2_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 18, self.g2_s2_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 18, self.g2_s2_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g2_s2_dict['maxMark'] == self.xl_G2_S2_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 19, self.g2_s2_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 19, self.g2_s2_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G2_S2_Min[loop] is None:
            self.xl_G2_S2_Min[loop] = 0
            if self.g2_s2_dict['minMark'] == self.xl_G2_S2_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 20, self.g2_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 20, self.g2_s2_dict['minMark'], self.style3)
        else:
            if round(-self.g2_s2_dict['minMark'], 2) == round(self.xl_G2_S2_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 20, -self.g2_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 20, -self.g2_s2_dict['minMark'], self.style3)
        # -----------------------------------------------------Group 3 -------------------------------------------------
        if self.g3_dict['groupName'] == self.xl_G3_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 21, self.g3_dict['groupName'], self.style8)
        else:
            self.ws.write(self.rowsize, 21, self.g3_dict['groupName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g3_dict['maxMark'] == self.xl_G3_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 22, self.g3_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 22, self.g3_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G3_Min[loop] is None:
            self.xl_G3_Min[loop] = 0
            if self.g3_dict['minMark'] == self.xl_G3_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 23, self.g3_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 23, self.g3_dict['minMark'], self.style3)
        else:
            if round(-self.g3_dict['minMark'], 2) == round(self.xl_G3_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 23, -self.g3_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 23, -self.g3_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g3_s1_dict['sectionName'] == self.xl_G3_S1_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 24, self.g3_s1_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 24, self.g3_s1_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g3_s1_dict['maxMark'] == self.xl_G3_S1_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 25, self.g3_s1_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 25, self.g3_s1_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G3_S1_Min[loop] is None:
            self.xl_G3_S1_Min[loop] = 0
            if self.g3_s1_dict['minMark'] == self.xl_G3_S1_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 26, self.g3_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 26, self.g3_s1_dict['minMark'], self.style3)
        else:
            if round(-self.g3_s1_dict['minMark'], 2) == round(self.xl_G3_S1_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 26, -self.g3_s1_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 26, -self.g3_s1_dict['minMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g3_s2_dict['sectionName'] == self.xl_G3_S2_Name[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 27, self.g3_s2_dict['sectionName'], self.style8)
        else:
            self.ws.write(self.rowsize, 27, self.g3_s2_dict['sectionName'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.g3_s2_dict['maxMark'] == self.xl_G3_S2_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 28, self.g3_s2_dict['maxMark'], self.style8)
        else:
            self.ws.write(self.rowsize, 28, self.g3_s2_dict['maxMark'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.xl_G3_S2_Min[loop] is None:
            self.xl_G3_S2_Min[loop] = 0
            if self.g3_s2_dict['minMark'] == self.xl_G3_S2_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 29, self.g3_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 29, self.g3_s2_dict['minMark'], self.style3)
        else:
            if (-self.g3_s2_dict['minMark']) == self.xl_G3_S2_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 29, -self.g3_s2_dict['minMark'], self.style8)
            else:
                self.ws.write(self.rowsize, 29, -self.g3_s2_dict['minMark'], self.style3)
        # ---------------------------------------------Test-------- ----------------------------------------------------
        if self.xl_Test_Min[loop] is None:
            self.xl_Test_Min[loop] = 0
            if self.assessment_dict['minMarks'] == self.xl_Test_Min[loop]:
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 30, self.assessment_dict['minMarks'], self.style8)
            else:
                self.ws.write(self.rowsize, 30, self.assessment_dict['minMarks'], self.style3)
        else:
            if round(-self.assessment_dict['minMarks'], 2) == round(self.xl_Test_Min[loop], 2):
                self.Actual_Success_case_each.append('Pass')
                self.ws.write(self.rowsize, 30, -self.assessment_dict['minMarks'], self.style8)
            else:
                self.ws.write(self.rowsize, 30, -self.assessment_dict['minMarks'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.assessment_dict['maxMarks'] == self.xl_Test_Max[loop]:
            self.Actual_Success_case_each.append('Pass')
            self.ws.write(self.rowsize, 31, self.assessment_dict['maxMarks'], self.style8)
        else:
            self.ws.write(self.rowsize, 31, self.assessment_dict['maxMarks'], self.style3)
        # ----------------------------------------------------- --------------------------------------------------------
        if self.Expected_success_cases_each == self.Actual_Success_case_each:
            self.success_case_01 = 'Pass'
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # ------------------------------------ OutPut File save --------------------------------------------------------
        self.rowsize += 1
        Object.wb_Result.save(output_paths.outputpaths['min_max_Score'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Min_Max_Score', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
            print("Overall Status Pass")
        else:
            self.ws.write(0, 1, 'Fail', self.style25)
            print("Overall Status Fail")

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        # ---------------------------- OutPut File save with Overall Status --------------------------------------------
        Object.wb_Result.save(output_paths.outputpaths['min_max_Score'])
Object = Min_Max_Verification()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_testId)
print("Number Of Rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration ::", looping)
        Object.api_call(looping)
        Object.output_report(looping)
        # ----------------- Make Dictionaries clear for each loop ------------------------------------------------------
        Object.success_case_01 = {}
        Object.headers = {}
        Object.Actual_Success_case_each = []
# ---------------------------- Call this function at last --------------------------------------------------------------
Object.overall_status()