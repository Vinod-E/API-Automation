import pymysql
import time
import json
import requests
import xlwt
import datetime
import xlrd


class UploadCandidate:
    def __init__(self):
        self.start_time = str(datetime.datetime.now())

        # ------------------------
        # KMBDEMO LOGIN APPLICATION
        # ------------------------
        self.header = {"content-type": "application/json"}
        self.login_request = {"LoginName": 'admin',
                              "Password": 'admin@123',
                              "TenantAlias": "staffingautomation",
                              "UserName": 'admin'}

        login_api = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                  headers=self.header,
                                  data=json.dumps(self.login_request),
                                  verify=False)
        self.response = login_api.json()
        self.get_token = {"content-type": "application/json",
                          "X-AUTH-TOKEN": self.response.get("Token")}
        self.var = None
        time.sleep(1)
        resp_dict = json.loads(login_api.content)
        self.status = resp_dict['status']
        if self.status == 'OK':
            self.login = 'OK'
            print("Login successfully")
            print("Status is", self.status)
            time.sleep(1)
        else:
            self.login = 'KO'
            print("Failed to login")
            print("Status is", self.status)

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_ThirdPartyId = []
        self.xl_CandidateName = []
        self.xl_Email = []
        self.xl_Mobile1 = []
        self.xl_Gender = []
        self.xl_DateOfBirth = []
        self.xl_BusinessUnit = []
        self.xl_Department = []
        self.xl_OfferedRole = []
        self.xl_Spoc = []
        self.xl_PassportNo = []
        self.xl_Level = []
        self.xl_PanNo = []
        self.xl_OfferedLocation = []
        self.xl_ExpectedDOJ = []
        self.xl_TentativeDOJ = []
        self.xl_Integer1 = []
        self.xl_Integer2 = []
        self.xl_Integer3 = []
        self.xl_Integer10 = []
        self.xl_Integer12 = []
        self.xl_Text1 = []
        self.xl_Text2 = []
        self.xl_Text10 = []
        self.xl_Text11 = []
        self.xl_Text12 = []
        self.xl_Expected_msg = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 62)))

        # -------------------------------------------------------
        # Styles for Excel sheet Row, Column, Text - color, Font
        # -------------------------------------------------------
        self.__style0 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                    'font: name Arial, color black, bold on;')
        self.__style1 = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'
                                    'font: name Arial, color black, bold off;')
        self.__style2 = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                                    'font: name Arial, color yellow, bold on;')
        self.__style3 = xlwt.easyxf('font: name Arial, color red, bold on')
        self.__style4 = xlwt.easyxf('pattern: pattern solid, fore_colour indigo;'
                                    'font: name Arial, color gold, bold on;')
        self.__style5 = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;'
                                    'font: name Arial, color brown, bold on;')
        self.__style6 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.__style7 = xlwt.easyxf('font: name Arial, color orange, bold on')
        self.__style8 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.__style9 = xlwt.easyxf('font: name Arial, color green, bold on')

        # -------------------------------------
        # Excel sheet write for Output results
        # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Upload_Candidates')
        self.rowsize = 2
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparison', 'Remarks', 'current_status', 'InitiatedOn', 'Id', 'candidateUserId', 'CurrentActivity',
                        'ThirdPartyId',
                        'Name', 'Email', 'Mobile1', 'Gender', 'DateOfBirth', 'Spoc', 'OfferedBUId', 'OwingDepartmentId',
                        'OfferedDesignationId', 'PassportNo', 'LevelId', 'PanNo', 'LocationOfferedId', 'ExpectedDOJ',
                        'TentativeDoj', 'Integer1', 'Integer2', 'Integer3', 'Integer10', 'Integer12', 'Text1', 'Text2',
                        'Text10', 'Text11', 'Text12']

        for headers in excelheaders:
            if headers in ['Comparison', 'Remarks', 'current_status', 'InitiatedOn', 'Id', 'CandidateUserId', 'CurrentActivity']:
                self.ws.write(1, index, headers, self.__style2)
            else:
                self.ws.write(1, index, headers, self.__style0)
            index += 1

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------

        self.wb = xlrd.open_workbook(
            '/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
            '/UploadExcel07_AUTOMATION.xls')
        self.sheet1 = self.wb.sheet_by_index(0)
        for i in range(1, self.sheet1.nrows):
            number = i
            rows = self.sheet1.row_values(number)

            # ----------------------------------------------------------
            # Candidate Personal details
            # ----------------------------------------------------------

            if not rows[0]:
                self.xl_ThirdPartyId.append(None)
            else:
                self.xl_ThirdPartyId.append(str(rows[0]))

            if not rows[1]:
                self.xl_CandidateName.append(None)
            else:
                self.xl_CandidateName.append(str(rows[1]))

            if not rows[2]:
                self.xl_Email.append(None)
            else:
                self.xl_Email.append(rows[2])

            if not rows[3]:
                self.xl_Mobile1.append(None)
            else:
                self.xl_Mobile1.append(str(rows[3]))

            if not rows[4]:
                self.xl_Gender.append(None)
            else:
                self.xl_Gender.append(rows[4])

            if not rows[5]:
                self.xl_DateOfBirth.append(None)

            else:

                DateOfBirth = self.sheet1.cell_value(rowx=(i), colx=5)
                self.DateOfBirth = datetime.datetime(*xlrd.xldate_as_tuple(DateOfBirth, self.wb.datemode))
                self.DateOfBirth = self.DateOfBirth.strftime("%d-%m-%Y")
                self.xl_DateOfBirth.append(self.DateOfBirth)
                print(self.xl_DateOfBirth)

            if not rows[6]:
                self.xl_Spoc.append(None)
            else:
                self.xl_Spoc.append(rows[6])

            if not rows[7]:
                self.xl_BusinessUnit.append(None)
            else:
                self.xl_BusinessUnit.append(rows[7])

            if not rows[8]:
                self.xl_Department.append(None)
            else:
                self.xl_Department.append(rows[8])

            if not rows[9]:
                self.xl_OfferedRole.append(None)
            else:
                self.xl_OfferedRole.append(rows[9])

            if not rows[10]:
                self.xl_PassportNo.append(None)
            else:
                self.xl_PassportNo.append(rows[10])

            if not rows[11]:
                self.xl_Level.append(None)
            else:
                self.xl_Level.append(rows[11])

            if not rows[12]:
                self.xl_PanNo.append(None)
            else:
                self.xl_PanNo.append(rows[12])

            if not rows[13]:
                self.xl_OfferedLocation.append(None)
            else:
                self.xl_OfferedLocation.append(rows[13])

            if not rows[14]:
                self.xl_ExpectedDOJ.append(None)

            else:
                ExpectedDOJ = self.sheet1.cell_value(rowx=(i), colx=14)
                self.ExpectedDOJ = datetime.datetime(*xlrd.xldate_as_tuple(ExpectedDOJ, self.wb.datemode))
                self.ExpectedDOJ = self.ExpectedDOJ.strftime("%d-%m-%Y")
                self.xl_ExpectedDOJ.append(self.ExpectedDOJ)
                print(self.xl_ExpectedDOJ)

            if not rows[15]:
                self.xl_TentativeDOJ.append(None)

            else:
                TentativeDOJ = self.sheet1.cell_value(rowx=(i), colx=15)
                self.TentativeDOJ = datetime.datetime(*xlrd.xldate_as_tuple(TentativeDOJ, self.wb.datemode))
                self.TentativeDOJ = self.TentativeDOJ.strftime("%d-%m-%Y")
                self.xl_TentativeDOJ.append(self.TentativeDOJ)
                print(self.xl_TentativeDOJ)

            if not rows[16]:
                self.xl_Integer1.append(None)
            else:
                self.xl_Integer1.append(rows[16])

            if not rows[17]:
                self.xl_Integer2.append(None)
            else:
                self.xl_Integer2.append(rows[17])

            if not rows[18]:
                self.xl_Integer3.append(None)
            else:
                self.xl_Integer3.append(rows[18])

            if not rows[19]:
                self.xl_Integer10.append(None)
            else:
                self.xl_Integer10.append(rows[19])

            if not rows[20]:
                self.xl_Integer12.append(None)
            else:
                self.xl_Integer12.append(rows[20])

            if not rows[21]:
                self.xl_Text1.append(None)
            else:
                self.xl_Text1.append(rows[21])

            if not rows[22]:
                self.xl_Text2.append(None)
            else:
                self.xl_Text2.append(rows[22])

            if not rows[23]:
                self.xl_Text10.append(None)
            else:
                self.xl_Text10.append(rows[23])

            if not rows[24]:
                self.xl_Text11.append(None)
            else:
                self.xl_Text11.append(rows[24])

            if not rows[25]:
                self.xl_Text12.append(None)
            else:
                self.xl_Text12.append(rows[25])

            if not rows[26]:
                self.xl_Expected_msg.append(None)
            else:
                self.xl_Expected_msg.append(rows[26])

    # --------------- Candidate create request ------------------
    def api_call_bulk_upload(self, iteration_count):
        try:

            conn = pymysql.connect(host='35.154.36.218',
                                           database='appserver_core',
                                           user='qauser',
                                           password='qauser')
            cur1 = conn.cursor()
            self.total = 1
            self.cand_list = []
            print(cur1)
            self.data = {"type": "offerImport",
                         "candidateList": [{"ThirdPartyID": self.xl_ThirdPartyId[iteration_count],
                                            "Name": self.xl_CandidateName[iteration_count],
                                            "Email1": self.xl_Email[iteration_count],
                                            "Mobile1": self.xl_Mobile1[iteration_count],
                                            "Gender": self.xl_Gender[iteration_count],
                                            "DateOfBirth": self.xl_DateOfBirth[iteration_count],

                                            "PassportNo": self.xl_PassportNo[iteration_count],
                                            "PanNo": self.xl_PanNo[iteration_count],
                                            "ExpectedDOJ": self.xl_ExpectedDOJ[iteration_count],
                                            "TentativeDOJ": self.xl_TentativeDOJ[iteration_count],
                                            "Integer1": self.xl_Integer1[iteration_count],
                                            "Integer2": self.xl_Integer2[iteration_count],
                                            "Integer3": self.xl_Integer3[iteration_count],
                                            "Integer10": self.xl_Integer10[iteration_count],
                                            "Integer12": self.xl_Integer12[iteration_count],
                                            "Text1": self.xl_Text1[iteration_count],
                                            "Text2": self.xl_Text2[iteration_count],
                                            "Text10": self.xl_Text10[iteration_count],
                                            "Text11": self.xl_Text11[iteration_count],
                                            "Text12": self.xl_Text12[iteration_count],
                                            "SPOCId": self.xl_Spoc[iteration_count],
                                            "hasError": False,
                                            "duplicateObj": [],
                                            "isDuplicationCkd": False,
                                            "isDuplicate": False,
                                            "LocationOfferedId": self.xl_OfferedLocation[iteration_count],
                                            "OfferedDesignationId": self.xl_OfferedRole[iteration_count],
                                            "OfferedBUId": self.xl_BusinessUnit[iteration_count],
                                            "LevelId": self.xl_Level[iteration_count],
                                            "OwningDepartmentId": self.xl_Department[
                                                iteration_count]

                                            }]}

            Create_candidate = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/bulkimport",
                                             headers=self.get_token,
                                             data=json.dumps(self.data, default=str), verify=False)
            print("Below is Original Data")
            print(self.data)
            print("Below1 is Original Data")
            create_candidate_response_resp_dict = json.loads(Create_candidate.content)

            print(create_candidate_response_resp_dict)
            time.sleep(1)
            self.status = create_candidate_response_resp_dict['status']
            data = create_candidate_response_resp_dict.get('data')
            a = data.get('importresult')
            failed = a.get('FailedList')
            if failed and len(failed):
                for failed_data in a.get('FailedList'):
                    if failed_data.get('Error'):
                        # self.ws.write(self.rowsize, 1, failed_data.get('Error'), self.__style9)
                        # self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                        if failed_data:
                            self.ws.write(self.rowsize, 1, failed_data.get('Error'), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 1, 'String mismatch', self.__style3)
                        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                        self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count], self.__style9)

                        self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style9)
                        self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style9)

                        # self.ws.write(self.rowsize, 32, self.xl_TrueFalse1[Iteration_Count], self.__style1)
                        # self.ws.write(self.rowsize, 33, self.xl_TrueFalse5[Iteration_Count], self.__style1)
                        self.rowsize += 1  # Row increment
                        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                        if failed_data:
                            if self.xl_Expected_msg[Iteration_Count] in failed_data:
                                self.ws.write(self.rowsize, 1, failed_data.get('Error'), self.__style9)
                            self.ws.write(self.rowsize, 2, 'pass', self.__style9)
                        self.rowsize += 1  # Row increment
                        ob.wb_Result.save(
                            '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                            '/Upload_Candidates.xls')


                    else:
                        for failed_data in a.get('FailedList'):
                            print(failed_data.values()[0])
                            self.ws.write(self.rowsize, 1, failed_data.values()[0], self.__style9)
                            self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                            self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style9)
                            self.rowsize += 1  # Row increment
                            self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

                            if failed_data:
                                if self.xl_Expected_msg[Iteration_Count] in failed_data.values()[0]:
                                    self.ws.write(self.rowsize, 1, failed_data.values()[0], self.__style9)
                                self.ws.write(self.rowsize, 2, 'pass', self.__style9)

                            else:
                                self.ws.write(self.rowsize, 1, 'String mismatch', self.__style3)
                            # self.rowsize += 1  # Row increment
                            # self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                            self.rowsize += 1  # Row increment
                            ob.wb_Result.save(
                                '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                                '/UploadCandidates.xls')

            duplicate = a.get('DuplicateLIst')
            if len(duplicate):
                for Values in duplicate:
                    for duplicate_data in Values.get('duplicateValues'):
                        print(duplicate_data)
                        print(Values)
                        if duplicate_data == "Email1":
                            email1 = Values.get("originalCandidateInfo").get("Email1")
                            b = "select id from candidates where  hp_dec(email1) like '%s'" % email1
                            cur1.execute(b)
                            candidate_id = cur1.fetchone()
                            self.cand_list.append(candidate_id[0])
                            self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                            self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style9)
                            self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style9)
                            # self.ws.write(self.rowsize, 32, self.xl_TrueFalse1[Iteration_Count], self.__style1)
                            # self.ws.write(self.rowsize, 33, self.xl_TrueFalse5[Iteration_Count], self.__style1)
                            self.ws.write(self.rowsize, 1, ('Email1 ', str(candidate_id)),
                                          self.__style3)
                            self.rowsize += 1  # Row increment
                            self.ws.write(self.rowsize, self.col, 'Output', self.__style9)

                            if duplicate_data:
                                if self.xl_Expected_msg[Iteration_Count] == duplicate_data:
                                    self.ws.write(self.rowsize, 1, Values.get('duplicateValues'),
                                                  self.__style9)
                                self.ws.write(self.rowsize, 2, 'pass', self.__style9)

                            # self.rowsize += 1  # Row increment
                            # # self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                            # self.ws.write(self.rowsize, 1, ('Email ', str(candidate_id)),
                            #               self.__style3)

                            self.rowsize += 1  # Row increment
                            ob.wb_Result.save(
                                '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                                '/UploadCandidates.xls')
                        else:
                            for value in a.get('DuplicateLIst'):
                                for duplicate_failed_data in value.get('duplicateValues'):
                                    if duplicate_failed_data.get('Mobile1'):
                                        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                                        self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count],
                                                      self.__style1)
                                        self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style1)
                                        self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style1)
                                        # self.ws.write(self.rowsize, 32, self.xl_TrueFalse1[Iteration_Count],
                                        #               self.__style1)
                                        # self.ws.write(self.rowsize, 33, self.xl_TrueFalse5[Iteration_Count],
                                        # self.__style1)
                                        # self.rowsize += 1  # Row increment
                                        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                                        self.ws.write(self.rowsize, 1, ('Candidate is duplicate with', ('Mobile1')),
                                                      self.__style3)
                                        if duplicate_data:
                                            if self.xl_Expected_msg[Iteration_Count] == duplicate_failed_data:
                                                self.ws.write(self.rowsize, 1, duplicate_failed_data.get('Mobile1'),
                                                              self.__style9)
                                            self.ws.write(self.rowsize, 2, 'pass', self.__style9)
                                        self.rowsize += 1  # Row increment
                                        ob.wb_Result.save(
                                            '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu'
                                            '/Upload_candidates/UploadCandidates.xls')

            success = a.get('successList')
            if len(success):
                for success_list in a.get('successList'):
                    d = "select id from candidates where  hp_dec(email1) like '%s'" % success_list
                    cur1.execute(d)
                    candidate_id = cur1.fetchone()
                    self.cand_list.append(candidate_id[0])

                    header = {"Content-type": "application/json"}
                    data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "staffingautomation",
                            "UserName": "admin"}
                    response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                             data=json.dumps(data), verify=False)
                    Token = response.json().get("Token")
                    headers = {"content-type": "application/json", "X-AUTH-TOKEN": Token}
                    for cand in self.cand_list:
                        candidate_get_by_id = requests.post(
                            "https://amsin.hirepro.in/py/pofu/api/v1/get-candidate-by-id/{}/".format(cand),
                            headers=headers,
                            verify=False)
                        self.resp_dict = json.loads(candidate_get_by_id.content)
                        print(self.resp_dict)

                        # self.message =

                        candidate = self.resp_dict.get('Candidate')
                        self.ID = candidate.get('Id')
                        self.InitiatedOn = candidate.get('IntitatedOn')
                        self.CandidateUserId = candidate.get('CandidateUserId')
                        self.CurrentActivity = candidate.get('CurrentActivity')
                        self.api_candidate_ThirdPartyId = self.resp_dict["Candidate"]["ThirdPartyId"]
                        self.api_candidate_CandidateName = self.resp_dict["Candidate"]["CandidateName"]
                        self.api_candidate_Email = self.resp_dict["Candidate"]["Email"]
                        self.api_candidate_Mobile1 = self.resp_dict["Candidate"]["ContactNumber"]
                        print(self.api_candidate_Mobile1)
                        self.api_candidate_Gender = self.resp_dict["Candidate"]["Gender"]
                        self.api_candidate_DateOfBirth = '-'.join(
                            (self.resp_dict["Candidate"]["DateOfBirth"].replace('T00:00:00', "")).split('-')[::-1])
                        self.api_candidate_Spoc = (self.resp_dict["Candidate"]["Spoc"])
                        self.api_candidate_BusinessUnit = self.resp_dict["Candidate"]["BusinessUnitId"]
                        self.api_candidate_Department = self.resp_dict["Candidate"]["DepartmentId"]
                        self.api_candidate_OfferedRole = self.resp_dict["Candidate"]["OfferedRoleId"]

                        self.api_candidate_PassportNo = self.resp_dict["Candidate"]["PassportNo"]
                        self.api_candidate_Level = self.resp_dict["Candidate"]["LevelId"]
                        self.api_candidate_PanNo = self.resp_dict["Candidate"]["PanNo"]
                        self.api_candidate_OfferedLocation = self.resp_dict["Candidate"]["OfferedLocation"]
                        self.api_candidate_ExpectedDOJ = '-'.join(
                            (self.resp_dict["Candidate"]["ExpectedJoiningDate"].replace('T00:00:00', "")).split('-')[
                            ::-1])
                        self.api_candidate_TentativeDOJ = '-'.join(
                            (self.resp_dict["Candidate"]["TentativeDOJ"].replace('T00:00:00', "")).split('-')[::-1])
                        self.api_candidate_Integer1 = self.resp_dict["Candidate"]["Integer1"]
                        self.api_candidate_Integer2 = self.resp_dict["Candidate"]["Integer2"]
                        self.api_candidate_Integer3 = self.resp_dict["Candidate"]["Integer3"]
                        self.api_candidate_Integer10 = self.resp_dict["Candidate"]["Integer10"]
                        self.api_candidate_Integer12 = self.resp_dict["Candidate"]["Integer12"]

                        self.api_candidate_Text1 = self.resp_dict["Candidate"]["Text1"]
                        self.api_candidate_Text2 = self.resp_dict["Candidate"]["Text2"]
                        self.api_candidate_Text10 = self.resp_dict["Candidate"]["Text10"]
                        self.api_candidate_Text11 = self.resp_dict["Candidate"]["Text11"]
                        self.api_candidate_Text12 = self.resp_dict["Candidate"]["Text12"]
                        # self.api_candidate_TrueFalse1 = self.resp_dict["Candidate"]["TrueFalse1"]
                        # self.api_candidate_TrueFalse5 = self.resp_dict["Candidate"]["TrueFalse5"]
                        # =============================================================================================
                        #       Writing Excel Value to the output Excel
                        # =============================================================================================

                        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                        self.ws.write(self.rowsize, 1, 'Excel Input', self.__style9)
                        if self.xl_ThirdPartyId[Iteration_Count]:
                            self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 7, 'Empty', self.__style9)
                        if self.xl_CandidateName[Iteration_Count]:
                            self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 8, 'Empty', self.__style9)
                        if self.xl_Email[Iteration_Count]:
                            self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style9)

                        else:
                            self.ws.write(self.rowsize, 9, 'Empty', self.__style9)
                        if self.xl_Mobile1[Iteration_Count]:
                            self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 10, 'Empty', self.__style9)
                        if self.xl_Gender[Iteration_Count]:
                            self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 11, 'Empty', self.__style9)
                        if self.xl_DateOfBirth[Iteration_Count]:
                            self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 12, 'Empty', self.__style9)
                        if self.xl_Spoc[Iteration_Count]:
                            self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 13, 'Empty', self.__style9)
                        if self.xl_BusinessUnit[Iteration_Count]:
                            self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 14, 'Empty', self.__style9)
                        if self.xl_Department[Iteration_Count]:
                            self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 15, 'Empty', self.__style9)
                        if self.xl_OfferedRole[Iteration_Count]:
                            self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 16, 'Empty', self.__style9)
                        if self.xl_PassportNo[Iteration_Count]:
                            self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 17, 'Empty', self.__style9)
                        if self.xl_Level[Iteration_Count]:
                            self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 18, 'Empty', self.__style9)
                        if self.xl_PanNo[Iteration_Count]:
                            self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 19, 'Empty', self.__style9)
                        if self.xl_OfferedLocation[Iteration_Count]:
                            self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 20, 'Empty', self.__style9)
                        if self.xl_ExpectedDOJ[Iteration_Count]:
                            self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 21, 'Empty', self.__style9)
                        if self.xl_TentativeDOJ[Iteration_Count]:
                            self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 22, 'Empty', self.__style9)
                        if self.xl_Integer1[Iteration_Count]:
                            self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 23, 'Empty', self.__style9)
                        if self.xl_Integer2[Iteration_Count]:
                            self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 24, 'Empty', self.__style9)
                        if self.xl_Integer3[Iteration_Count]:
                            self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 25, 'Empty', self.__style9)
                        if self.xl_Integer10[Iteration_Count]:
                            self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 26, 'Empty', self.__style9)
                        if self.xl_Integer12[Iteration_Count]:
                            self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 27, 'Empty', self.__style9)
                        if self.xl_Text1[Iteration_Count]:
                            self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 28, 'Empty', self.__style9)
                        if self.xl_Text2[Iteration_Count]:
                            self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 29, 'Empty', self.__style9)
                        if self.xl_Text10[Iteration_Count]:
                            self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 30, 'Empty', self.__style9)
                        if self.xl_Text11[Iteration_Count]:
                            self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 31, 'Empty', self.__style9)
                        if self.xl_Text12[Iteration_Count]:
                            self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style9)
                        else:
                            self.ws.write(self.rowsize, 32, 'Empty', self.__style9)
                        # self.ws.write(self.rowsize, 32, self.xl_TrueFalse1[Iteration_Count], self.__style1)
                        # self.ws.write(self.rowsize, 33, self.xl_TrueFalse5[Iteration_Count], self.__style1)

                        # ----------------------
                        #   Write Output Data
                        # ----------------------
                        self.rowsize += 1  # Row increment
                        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

                        if self.cand_list:
                            self.ws.write(self.rowsize, 1, 'Uploaded successfully', self.__style9)
                        else:
                            self.ws.write(self.rowsize, 1, 'Not able to Upload', self.__style3)
                        # -----------------------------------------------------------------------------------------------
                        self.ws.write(self.rowsize, 3, self.InitiatedOn)
                        self.ws.write(self.rowsize, 4, self.ID)
                        self.ws.write(self.rowsize, 5, self.CandidateUserId)
                        self.ws.write(self.rowsize, 6, self.CurrentActivity)

                        # ------------------------------------------------------------------
                        # Comparing API Data with Excel Data and Printing into Output Excel
                        # ------------------------------------------------------------------

                        if self.api_candidate_ThirdPartyId:
                            if str(self.xl_ThirdPartyId[Iteration_Count]) == str(self.api_candidate_ThirdPartyId):
                                self.ws.write(self.rowsize, 7, str(self.api_candidate_ThirdPartyId), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 7, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 7, 'Empty', self.__style9)
                        # ---------------------------------------------------------------------------
                        #  ThirdPartyId
                        # ---------------------------------------------------------------------------

                        if self.api_candidate_CandidateName:
                            if self.xl_CandidateName[Iteration_Count] == str(self.api_candidate_CandidateName):
                                self.ws.write(self.rowsize, 8, str(self.api_candidate_CandidateName), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 8, str(self.api_candidate_CandidateName), self.__style7)
                        else:
                            self.ws.write(self.rowsize, 8, 'Empty', self.__style9)
                        # --------------------------------------------------------------------------
                        #   Email
                        # --------------------------------------------------------------------------
                        if self.api_candidate_Email:
                            if self.xl_Email[Iteration_Count] == str(self.api_candidate_Email):
                                self.ws.write(self.rowsize, 9, str(self.api_candidate_Email), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 9, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 9, 'Empty', self.__style9)
                        # -----------------------------------------------------------------------------
                        #   Mobile1
                        # -----------------------------------------------------------------------------

                        if self.api_candidate_Mobile1:
                            if str(self.xl_Mobile1[Iteration_Count]) == str(self.api_candidate_Mobile1):
                                self.ws.write(self.rowsize, 10, str(self.api_candidate_Mobile1), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 10, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 10, 'Empty', self.__style9)

                        if self.api_candidate_Gender:
                            if self.xl_Gender[Iteration_Count] == (self.api_candidate_Gender):
                                self.ws.write(self.rowsize, 11, (self.api_candidate_Gender), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 11, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 11, 'Empty', self.__style9)

                        if self.api_candidate_DateOfBirth:
                            if str(self.xl_DateOfBirth[Iteration_Count]) == str(self.api_candidate_DateOfBirth):
                                self.ws.write(self.rowsize, 12, str(self.api_candidate_DateOfBirth), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 12, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 12, 'Empty', self.__style9)

                        if self.api_candidate_Spoc:
                            if self.xl_Spoc[Iteration_Count] == self.api_candidate_Spoc:
                                self.ws.write(self.rowsize, 13, self.api_candidate_Spoc, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 13, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 13, 'Empty', self.__style9)

                        if self.api_candidate_BusinessUnit:
                            if self.xl_BusinessUnit[Iteration_Count] == self.api_candidate_BusinessUnit:
                                self.ws.write(self.rowsize, 14, self.api_candidate_BusinessUnit, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 14, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 14, 'Empty', self.__style9)

                        if self.api_candidate_Department:
                            if self.xl_Department[Iteration_Count] == self.api_candidate_Department:
                                self.ws.write(self.rowsize, 15, self.api_candidate_Department, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 15, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 15, 'Empty', self.__style9)

                        if self.api_candidate_OfferedRole:
                            if self.xl_OfferedRole[Iteration_Count] == self.api_candidate_OfferedRole:
                                self.ws.write(self.rowsize, 16, self.api_candidate_OfferedRole, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 16, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 16, 'Empty', self.__style9)

                        if self.api_candidate_PassportNo:
                            if self.xl_PassportNo[Iteration_Count] == str(self.api_candidate_PassportNo):
                                self.ws.write(self.rowsize, 17, str(self.api_candidate_PassportNo), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 17, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 17, 'Empty', self.__style9)

                        if self.api_candidate_Level:
                            if self.xl_Level[Iteration_Count] == self.api_candidate_Level:
                                self.ws.write(self.rowsize, 18, self.api_candidate_Level, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 18, 'Fail', self.__style9)
                        else:
                            self.ws.write(self.rowsize, 18, 'Empty', self.__style9)

                        if self.api_candidate_PanNo:
                            if self.xl_PanNo[Iteration_Count] == str(self.api_candidate_PanNo):
                                self.ws.write(self.rowsize, 19, str(self.api_candidate_PanNo), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 19, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 19, 'Empty', self.__style9)

                        if self.api_candidate_OfferedLocation:
                            if self.xl_OfferedLocation[Iteration_Count] == self.api_candidate_OfferedLocation:
                                self.ws.write(self.rowsize, 20, self.api_candidate_OfferedLocation, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 20, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 20, 'Empty', self.__style9)

                        if self.api_candidate_ExpectedDOJ:
                            if self.xl_ExpectedDOJ[Iteration_Count] == str(self.api_candidate_ExpectedDOJ):
                                self.ws.write(self.rowsize, 21, str(self.api_candidate_ExpectedDOJ), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 21, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 21, 'Empty', self.__style9)

                        if self.api_candidate_TentativeDOJ:
                            if self.xl_TentativeDOJ[Iteration_Count] == str(self.api_candidate_TentativeDOJ):
                                self.ws.write(self.rowsize, 22, str(self.api_candidate_TentativeDOJ), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 22, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 22, 'Empty', self.__style9)

                        if self.api_candidate_Integer1:
                            if self.xl_Integer1[Iteration_Count] == self.api_candidate_Integer1:
                                self.ws.write(self.rowsize, 23, self.api_candidate_Integer1, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 23, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 23, 'Empty', self.__style9)

                        if self.api_candidate_Integer2:
                            if self.xl_Integer2[Iteration_Count] == self.api_candidate_Integer2:
                                self.ws.write(self.rowsize, 24, self.api_candidate_Integer2, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 24, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 24, 'Empty', self.__style9)

                        if self.api_candidate_Integer3:
                            if self.xl_Integer3[Iteration_Count] == self.api_candidate_Integer3:
                                self.ws.write(self.rowsize, 25, self.api_candidate_Integer3, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 25, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 25, 'Empty', self.__style9)

                        if self.api_candidate_Integer10:
                            if self.xl_Integer10[Iteration_Count] == self.api_candidate_Integer10:
                                self.ws.write(self.rowsize, 26, self.api_candidate_Integer10, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 26, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 26, 'Empty', self.__style9)

                        if self.api_candidate_Integer12:
                            if self.xl_Integer12[Iteration_Count] == self.api_candidate_Integer12:
                                self.ws.write(self.rowsize, 27, self.api_candidate_Integer12, self.__style9)
                            else:
                                self.ws.write(self.rowsize, 27, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 27, 'Empty', self.__style9)

                        if self.api_candidate_Text1:
                            if self.xl_Text1[Iteration_Count] == str(self.api_candidate_Text1):
                                self.ws.write(self.rowsize, 28, str(self.api_candidate_Text1), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 28, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 28, 'Empty', self.__style9)

                        if self.api_candidate_Text2:
                            if self.xl_Text2[Iteration_Count] == str(self.api_candidate_Text2):
                                self.ws.write(self.rowsize, 29, str(self.api_candidate_Text2), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 29, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 29, 'Empty', self.__style9)

                        if self.api_candidate_Text10:
                            if self.xl_Text10[Iteration_Count] == str(self.api_candidate_Text10):
                                self.ws.write(self.rowsize, 30, str(self.api_candidate_Text10), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 30, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 30, 'Empty', self.__style9)

                        if self.api_candidate_Text11:
                            if self.xl_Text11[Iteration_Count] == str(self.api_candidate_Text11):
                                self.ws.write(self.rowsize, 31, str(self.api_candidate_Text11), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 31, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 31, 'Empty', self.__style9)

                        if self.api_candidate_Text12:
                            if self.xl_Text12[Iteration_Count] == str(self.api_candidate_Text12):
                                self.ws.write(self.rowsize, 32, str(self.api_candidate_Text12), self.__style9)
                            else:
                                self.ws.write(self.rowsize, 32, 'Fail', self.__style7)
                        else:
                            self.ws.write(self.rowsize, 32, 'Empty', self.__style9)

                        if candidate_id:
                            self.ws.write(self.rowsize, 2, 'Pass', self.__style9)
                        else:
                            self.ws.write(self.rowsize, 2, 'Fail', self.__style7)

                        self.rowsize += 1  # Row increment
                        ob.wb_Result.save(
                            '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                            '/UploadCandidates.xls')

            probable = a.get('probableDuplicates')
            if len(probable):
                for Values in a.get('probableDuplicates'):
                    for Probable_duplicate in Values.get('duplicateValues'):
                        print(Probable_duplicate)
                        print(Values)

                        self.ws.write(self.rowsize, 1, 'Probable Duplicate - Input data', self.__style3)
                        # self.ws.write(self.rowsize, 1, ('Candidate is Probable duplicate with any of the excel data or other user is also trying to upload candidate'), self.__style3)
                        self.ws.write(self.rowsize, 7, self.xl_ThirdPartyId[Iteration_Count], self.__style1)

                        self.ws.write(self.rowsize, 8, self.xl_CandidateName[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 9, self.xl_Email[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 10, self.xl_Mobile1[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 11, self.xl_Gender[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 12, self.xl_DateOfBirth[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 13, self.xl_Spoc[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 14, self.xl_BusinessUnit[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 15, self.xl_Department[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 16, self.xl_OfferedRole[Iteration_Count], self.__style1)

                        self.ws.write(self.rowsize, 17, self.xl_PassportNo[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 18, self.xl_Level[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 19, self.xl_PanNo[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 20, self.xl_OfferedLocation[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 21, self.xl_ExpectedDOJ[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 22, self.xl_TentativeDOJ[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 23, self.xl_Integer1[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 24, self.xl_Integer2[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 25, self.xl_Integer3[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 26, self.xl_Integer10[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 27, self.xl_Integer12[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 28, self.xl_Text1[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 29, self.xl_Text2[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 30, self.xl_Text10[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 31, self.xl_Text11[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, 32, self.xl_Text12[Iteration_Count], self.__style1)
                        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                        self.rowsize += 1  # Row increment
                        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

                        if Probable_duplicate:
                            if self.xl_Expected_msg[Iteration_Count] == Values:
                                self.ws.write(self.rowsize, 1, Values.get('duplicateValues'),
                                              self.__style9)
                            self.ws.write(self.rowsize, 2, 'pass', self.__style9)


                        self.rowsize += 1  # Row increment
                        ob.wb_Result.save(
                            '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                            '/UploadCandidates.xls')

            cur = conn.cursor()
            conn.close()

        except Exception as e:
            print(e)
            print("DB Connection Error - Exception Block")

    def over_status(self):
        self.ws.write(0, 0, 'U_Candidates')
        if self.Expected_success_cases == self.Expected_success_cases:
            self.ws.write(0, 1, 'Pass')
        else:
            self.ws.write(0, 1, 'Fail')

        self.ws.write(0, 3, 'StartTime')
        self.ws.write(0, 4, self.start_time)
        ob.wb_Result.save('/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                          '/UploadCandidates.xls')


ob = UploadCandidate()
ob.excel_data()

total_len = len(ob.xl_ThirdPartyId)
for Iteration_Count in range(0, total_len):
    print("Iteration Count:- %s" % Iteration_Count)
    ob.api_call_bulk_upload(Iteration_Count)
ob.over_status()
    # ob.Db_connection(Iteration_Count)
    # ob.output_excel(Iteration_Count)
