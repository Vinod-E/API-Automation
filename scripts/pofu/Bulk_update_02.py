import pymysql
import time
import json
import requests
import xlwt
import datetime
import xlrd


class UpdateCandidate:
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
        self.xl_CandidateId = []
        self.xl_Name = []
        self.xl_Email = []
        self.xl_Mobile1 = []
        self.xl_Gender = []
        self.xl_DateOfBirth = []
        self.xl_BusinessUnit = []
        self.xl_Department = []
        self.xl_OfferedRole = []
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
        self.xl_Expected_msg = []
        self.Expected_success_cases = list(map(lambda x: 'FAIL', range(0, 1)))
        # self.xl_Text12 = []

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
        self.__style10 = xlwt.easyxf('font: name Arial, color Red, bold on;')

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
        excelheaders = ['Comparison', 'Current_status', 'Remark', 'CandidateId',
                        'Name', 'Email', 'Mobile1', 'Gender', 'DateOfBirth', 'OfferedBUId', 'OwingDepartmentId',
                        'OfferedDesignationId', 'PassportNo', 'LevelId', 'PanNo', 'LocationOfferedId', 'ExpectedDOJ',
                        'TentativeDoj', 'Integer1', 'Integer2', 'Integer3', 'Integer10', 'Integer12', 'Text1', 'Text2',
                        'Text10', 'Text11']

        for headers in excelheaders:
            if headers in ['Comparison', 'Current_status', 'Remark']:
                self.ws.write(1, index, headers, self.__style2)
            else:
                self.ws.write(1, index, headers, self.__style0)
            index += 1

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------

        self.wb = xlrd.open_workbook('/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                                     '/BulkUpdate_kmbdemo.xls')
        self.sheet1 = self.wb.sheet_by_index(0)
        for i in range(1, self.sheet1.nrows):
            number = i
            rows = self.sheet1.row_values(number)

            # ----------------------------------------------------------
            # Candidate Personal details
            # ----------------------------------------------------------

            if not rows[0]:
                self.xl_CandidateId.append(None)
            else:
                self.xl_CandidateId.append(int(rows[0]))

            if not rows[1]:
                self.xl_Name.append(None)
            else:
                self.xl_Name.append(str(rows[1]))

            if not rows[2]:
                self.xl_Email.append(None)
            else:
                self.xl_Email.append(rows[2])

            if not rows[3]:
                self.xl_Mobile1.append(None)
            else:
                self.xl_Mobile1.append(int(rows[3]))

            if not rows[4]:
                self.xl_Gender.append(None)
            else:
                self.xl_Gender.append(int(rows[4]))

            if not rows[5]:
                self.xl_DateOfBirth.append(None)

            else:

                DateOfBirth = self.sheet1.cell_value(rowx=(i), colx=5)
                self.DateOfBirth = datetime.datetime(*xlrd.xldate_as_tuple(DateOfBirth, self.wb.datemode))
                self.DateOfBirth = self.DateOfBirth.strftime("%d-%m-%Y")
                self.xl_DateOfBirth.append(self.DateOfBirth)
                print(self.xl_DateOfBirth)

            if not rows[6]:
                self.xl_BusinessUnit.append(None)
            else:
                self.xl_BusinessUnit.append(rows[6])

            if not rows[7]:
                self.xl_Department.append(None)
            else:
                self.xl_Department.append(rows[7])

            if not rows[8]:
                self.xl_OfferedRole.append(None)
            else:
                self.xl_OfferedRole.append(rows[8])

            if not rows[9]:
                self.xl_PassportNo.append(None)
            else:
                self.xl_PassportNo.append(rows[9])

            if not rows[10]:
                self.xl_Level.append(None)
            else:
                self.xl_Level.append(rows[10])

            if not rows[11]:
                self.xl_PanNo.append(None)
            else:
                self.xl_PanNo.append(rows[11])

            if not rows[12]:
                self.xl_OfferedLocation.append(None)
            else:
                self.xl_OfferedLocation.append(rows[12])

            if not rows[13]:
                self.xl_ExpectedDOJ.append(None)

            else:
                ExpectedDOJ = self.sheet1.cell_value(rowx=(i), colx=13)
                self.ExpectedDOJ = datetime.datetime(*xlrd.xldate_as_tuple(ExpectedDOJ, self.wb.datemode))
                self.ExpectedDOJ = self.ExpectedDOJ.strftime("%d-%m-%Y")
                self.xl_ExpectedDOJ.append(self.ExpectedDOJ)
                print(self.xl_ExpectedDOJ)

            if not rows[14]:
                self.xl_TentativeDOJ.append(None)

            else:
                TentativeDOJ = self.sheet1.cell_value(rowx=(i), colx=14)
                self.TentativeDOJ = datetime.datetime(*xlrd.xldate_as_tuple(TentativeDOJ, self.wb.datemode))
                self.TentativeDOJ = self.TentativeDOJ.strftime("%d-%m-%Y")
                self.xl_TentativeDOJ.append(self.TentativeDOJ)
                print(self.xl_TentativeDOJ)

            if not rows[15]:
                self.xl_Integer1.append(None)
            else:
                self.xl_Integer1.append(int(rows[15]))

            if not rows[16]:
                self.xl_Integer2.append(None)
            else:
                self.xl_Integer2.append(int(rows[16]))

            if not rows[17]:
                self.xl_Integer3.append(None)
            else:
                self.xl_Integer3.append(int(rows[17]))

            if not rows[18]:
                self.xl_Integer10.append(None)
            else:
                self.xl_Integer10.append(int(rows[18]))

            if not rows[19]:
                self.xl_Integer12.append(None)
            else:
                self.xl_Integer12.append(int(rows[19]))

            if not rows[20]:
                self.xl_Text1.append(None)
            else:
                self.xl_Text1.append(str(rows[20]))

            if not rows[21]:
                self.xl_Text2.append(None)
            else:
                self.xl_Text2.append(str(rows[21]))

            if not rows[22]:
                self.xl_Text10.append(None)
            else:
                self.xl_Text10.append(str(rows[22]))

            if not rows[23]:
                self.xl_Text11.append(None)
            else:
                self.xl_Text11.append(str(rows[23]))
            if not rows[24]:
                self.xl_Expected_msg.append(None)
            else:
                self.xl_Expected_msg.append(str(rows[24]))

                # if not rows[24]:
                #     self.xl_Text12.append(None)
                # else:
                #     self.xl_Text12.append(str(rows[24]))

                # -------------------------
                # Candidate create request
                # -------------------------

    def api_call_bulk_update(self, iteration_count):
        try:

            # conn = pymysql.connector.connect(host='35.154.36.218',
            # database='core1776',
            # user='qauser',
            # password='qauser')
            # cur1 = conn.cursor()
            # self.total = 1
            # self.cand_list = []
            # print cur1
            self.data = [{"CandidateId": self.xl_CandidateId[iteration_count],
                          "CandidateName": self.xl_Name[iteration_count],
                          "Email": self.xl_Email[iteration_count],
                          "Mobile": self.xl_Mobile1[iteration_count],
                          "Gender": self.xl_Gender[iteration_count],
                          "DateOfBirth": self.xl_DateOfBirth[iteration_count],
                          "OfferedBu": self.xl_BusinessUnit[iteration_count],
                          "OfferedDepartment": self.xl_Department[iteration_count],
                          "OfferedDesignation": self.xl_OfferedRole[iteration_count],
                          "PassportNo": self.xl_PassportNo[iteration_count],
                          "OfferedLevel": self.xl_Level[iteration_count],
                          "PanNo": self.xl_PanNo[iteration_count],
                          "OfferedLocation": self.xl_OfferedLocation[iteration_count],
                          "ExpectedDoj": self.xl_ExpectedDOJ[iteration_count],
                          "TentativeDoj": self.xl_TentativeDOJ[iteration_count],
                          "Integer1": self.xl_Integer1[iteration_count],
                          "Integer2": self.xl_Integer2[iteration_count],
                          "Integer3": self.xl_Integer3[iteration_count],
                          "Integer10": self.xl_Integer10[iteration_count],
                          "Integer12": self.xl_Integer12[iteration_count],
                          "Text1": self.xl_Text1[iteration_count],
                          "Text2": self.xl_Text2[iteration_count],
                          "Text10": self.xl_Text10[iteration_count],
                          "Text11": self.xl_Text11[iteration_count],
                          # "Text12": self.xl_Text2[iteration_count],
                          "hasError": False,
                          "duplicateObj": [],
                          "isDuplicationCkd": False,
                          "isDuplicate": False
                          }]
            update_candidate = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/bulk-update-candidate/",
                                             headers=self.get_token, data=json.dumps(self.data, default=str),
                                             verify=False)

            update_candidate_response_resp_dict = json.loads(update_candidate.content)

            time.sleep(1)
            update_candidate = update_candidate_response_resp_dict.get('failedCandidates')

            if len(update_candidate):
                for failed_data in update_candidate:
                    if failed_data.get('error'):
                        error = failed_data.get('error')
                        if isinstance(error, dict) is True:
                            if error.get("IsDuplicate") is True and error.get("DuplicateInfo"):
                                self.ws.write(self.rowsize, 1, 'Fail', self.__style10)
                                self.ws.write(self.rowsize, 2, 'Duplicate with Mobile', self.__style3)
                                self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                                self.ws.write(self.rowsize, 3, self.xl_CandidateId[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 4, self.xl_Name[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 5, self.xl_Email[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 6, self.xl_Mobile1[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 7, self.xl_Gender[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 8, self.xl_DateOfBirth[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 9, self.xl_BusinessUnit[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 10, self.xl_Department[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 11, self.xl_OfferedRole[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 12, self.xl_PassportNo[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 13, self.xl_Level[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 14, self.xl_PanNo[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 15, self.xl_OfferedLocation[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 16, self.xl_ExpectedDOJ[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 17, self.xl_TentativeDOJ[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 18, self.xl_Integer1[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 19, self.xl_Integer2[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 20, self.xl_Integer3[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 21, self.xl_Integer10[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 22, self.xl_Integer12[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 23, self.xl_Text1[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 24, self.xl_Text2[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 25, self.xl_Text10[Iteration_Count], self.__style2)
                                self.ws.write(self.rowsize, 26, self.xl_Text11[Iteration_Count], self.__style2)
                                # self.ws.write(self.rowsize, 26, self.xl_Text12[Iteration_Count], self.__style1)
                                self.rowsize += 1
                                # Row increment
                                self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                                self.rowsize += 1  # Row increment
                                ob.wb_Result.save(
                                    '/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                                    '/BulkUpdate_kmbdemo_output.xls')
                        else:
                            #self.ws.write(self.rowsize, 1, 'Fail', self.__style10)
                            self.ws.write(self.rowsize, 2, failed_data.get('error'), self.__style9)
                            self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                            self.ws.write(self.rowsize, 3, self.xl_CandidateId[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 4, self.xl_Name[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 5, self.xl_Email[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 6, self.xl_Mobile1[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 7, self.xl_Gender[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 8, self.xl_DateOfBirth[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 9, self.xl_BusinessUnit[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 10, self.xl_Department[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 11, self.xl_OfferedRole[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 12, self.xl_PassportNo[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 13, self.xl_Level[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 14, self.xl_PanNo[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 15, self.xl_OfferedLocation[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 16, self.xl_ExpectedDOJ[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 17, self.xl_TentativeDOJ[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 18, self.xl_Integer1[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 19, self.xl_Integer2[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 20, self.xl_Integer3[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 21, self.xl_Integer10[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 22, self.xl_Integer12[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 23, self.xl_Text1[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 24, self.xl_Text2[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 25, self.xl_Text10[Iteration_Count], self.__style2)
                            self.ws.write(self.rowsize, 26, self.xl_Text11[Iteration_Count], self.__style2)
                            # self.ws.write(self.rowsize, 26, self.xl_Text12[Iteration_Count], self.__style1)
                            self.rowsize += 1  # Row increment
                            self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                            if failed_data:
                                if self.xl_Expected_msg[Iteration_Count] in error:
                                    self.ws.write(self.rowsize, 2, failed_data.get('error'), self.__style9)
                                else:
                                    self.ws.write(self.rowsize, 2, failed_data.get('error'), self.__style7)
                                self.ws.write(self.rowsize, 1, 'pass', self.__style9)
                            self.rowsize += 1  # Row increment
                            ob.wb_Result.save(
                                '/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                                '/BulkUpdate_kmbdemo_output.xls')

                            print("email Is duplicate")

            # success = a.get('successList')
            self.cand_list = []
            update_candidate_01 = update_candidate_response_resp_dict.get('successCandidates')
            if len(update_candidate_01):
                for success_list in update_candidate_01:
                    # candidate_id = self.xl_CandidateId[0]


                    #  self.cand_list.append(aa)
                    candidate_id = success_list.get("CandidateId")
                    header = {"Content-type": "application/json"}
                    data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "staffingautomation",
                            "UserName": "admin"}
                    response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                             data=json.dumps(data), verify=False)
                    Token = response.json().get("Token")
                    headers = {"content-type": "application/json", "X-AUTH-TOKEN": Token}
                    candidate_get_by_id = requests.post(
                        "https://amsin.hirepro.in/py/pofu/api/v1/get-candidate-by-id/{}/".format(candidate_id),
                        headers=headers, verify=False)
                    self.resp_dict = json.loads(candidate_get_by_id.content)
                    print(self.resp_dict)

                    # self.message =

                    candidate = self.resp_dict.get('Candidate')
                    self.ID = candidate.get('Id')
                    # self.api_candidate_Id = self.resp_dict["Candidate"]["Id"]
                    self.api_candidate_CandidateName = self.resp_dict["Candidate"]["FirstName"]
                    self.api_candidate_Email = self.resp_dict["Candidate"]["Email"]
                    self.api_candidate_Mobile1 = self.resp_dict["Candidate"]["ContactNumber"]
                    print(self.api_candidate_Mobile1)
                    self.api_candidate_Gender = self.resp_dict["Candidate"]["Gender"]
                    self.api_candidate_DateOfBirth = datetime.datetime.strptime(
                        self.resp_dict["Candidate"]["DateOfBirth"], '%Y-%m-%dT%H:%M:%S').strftime('%d-%m-%Y')
                    self.api_candidate_BusinessUnit = self.resp_dict["Candidate"]["BusinessUnitId"]
                    self.api_candidate_Department = self.resp_dict["Candidate"]["DepartmentId"]
                    self.api_candidate_OfferedRole = self.resp_dict["Candidate"]["OfferedRoleId"]

                    self.api_candidate_PassportNo = self.resp_dict["Candidate"]["PassportNo"]
                    self.api_candidate_Level = self.resp_dict["Candidate"]["LevelId"]
                    self.api_candidate_PanNo = self.resp_dict["Candidate"]["PanNo"]
                    self.api_candidate_OfferedLocation = self.resp_dict["Candidate"]["OfferedLocation"]
                    self.api_candidate_ExpectedDOJ = datetime.datetime.strptime(
                        self.resp_dict["Candidate"]["ExpectedJoiningDate"],
                        '%Y-%m-%dT%H:%M:%S').strftime('%d-%m-%Y')
                    self.api_candidate_TentativeDOJ = datetime.datetime.strptime(
                        self.resp_dict["Candidate"]["TentativeDOJ"],
                        '%Y-%m-%dT%H:%M:%S').strftime('%d-%m-%Y')
                    self.api_candidate_Integer1 = self.resp_dict["Candidate"]["Integer1"]
                    self.api_candidate_Integer2 = self.resp_dict["Candidate"]["Integer2"]
                    self.api_candidate_Integer3 = self.resp_dict["Candidate"]["Integer3"]
                    self.api_candidate_Integer10 = self.resp_dict["Candidate"]["Integer10"]
                    self.api_candidate_Integer12 = self.resp_dict["Candidate"]["Integer12"]

                    self.api_candidate_Text1 = self.resp_dict["Candidate"]["Text1"]
                    self.api_candidate_Text2 = self.resp_dict["Candidate"]["Text2"]
                    self.api_candidate_Text10 = self.resp_dict["Candidate"]["Text10"]
                    self.api_candidate_Text11 = self.resp_dict["Candidate"]["Text11"]
                    # self.api_candidate_Text12 = self.resp_dict["Candidate"]["Text12"]
                    # =============================================================================================
                    #       Writing Excel Value to the output Excel
                    # =============================================================================================

                    self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                    # self.ws.write(self.rowsize, 1, 'Pass', self.__style9)
                    self.ws.write(self.rowsize, 2, 'Excel Input', self.__style9)
                    if self.xl_CandidateId[Iteration_Count]:
                        self.ws.write(self.rowsize, 3, self.xl_CandidateId[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 3, 'Empty', self.__style9)
                    if self.xl_Name[Iteration_Count]:
                        self.ws.write(self.rowsize, 4, self.xl_Name[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 4, 'Empty', self.__style9)
                    if self.xl_Email[Iteration_Count]:
                        self.ws.write(self.rowsize, 5, self.xl_Email[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 5, 'Empty', self.__style9)
                    if self.xl_Mobile1[Iteration_Count]:
                        self.ws.write(self.rowsize, 6, self.xl_Mobile1[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 6, 'Empty', self.__style9)
                    if self.xl_Gender[Iteration_Count]:
                        self.ws.write(self.rowsize, 7, self.xl_Gender[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 7, 'Empty', self.__style9)
                    if self.xl_DateOfBirth[Iteration_Count]:
                        self.ws.write(self.rowsize, 8, self.xl_DateOfBirth[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 8, 'Empty', self.__style9)
                    if self.xl_BusinessUnit[Iteration_Count]:
                        self.ws.write(self.rowsize, 9, self.xl_BusinessUnit[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 9, 'Empty', self.__style9)
                    if self.xl_Department[Iteration_Count]:
                        self.ws.write(self.rowsize, 10, self.xl_Department[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 10, 'Empty', self.__style9)
                    if self.xl_OfferedRole[Iteration_Count]:
                        self.ws.write(self.rowsize, 11, self.xl_OfferedRole[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 11, 'Empty', self.__style9)
                    if self.xl_PassportNo[Iteration_Count]:
                        self.ws.write(self.rowsize, 12, self.xl_PassportNo[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 12, 'Empty', self.__style9)
                    if self.xl_Level[Iteration_Count]:
                        self.ws.write(self.rowsize, 13, self.xl_Level[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 13, 'Empty', self.__style9)
                    if self.xl_PanNo[Iteration_Count]:
                        self.ws.write(self.rowsize, 14, self.xl_PanNo[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 14, 'Empty', self.__style9)
                    if self.xl_OfferedLocation[Iteration_Count]:
                        self.ws.write(self.rowsize, 15, self.xl_OfferedLocation[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 15, 'Empty', self.__style9)
                    if self.xl_ExpectedDOJ[Iteration_Count]:
                        self.ws.write(self.rowsize, 16, self.xl_ExpectedDOJ[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 16, 'Empty', self.__style9)
                    if self.xl_TentativeDOJ[Iteration_Count]:
                        self.ws.write(self.rowsize, 17, self.xl_TentativeDOJ[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 17, 'Empty', self.__style9)
                    if self.xl_Integer1[Iteration_Count]:
                        self.ws.write(self.rowsize, 18, self.xl_Integer1[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 18, 'Empty', self.__style9)
                    if self.xl_Integer2[Iteration_Count]:
                        self.ws.write(self.rowsize, 19, self.xl_Integer2[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 19, 'Empty', self.__style9)
                    if self.xl_Integer3[Iteration_Count]:
                        self.ws.write(self.rowsize, 20, self.xl_Integer3[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 20, 'Empty', self.__style9)
                    if self.xl_Integer10[Iteration_Count]:
                        self.ws.write(self.rowsize, 21, self.xl_Integer10[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 21, 'Empty', self.__style9)
                    if self.xl_Integer12[Iteration_Count]:
                        self.ws.write(self.rowsize, 22, self.xl_Integer12[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 22, 'Empty', self.__style9)
                    if self.xl_Text1[Iteration_Count]:
                        self.ws.write(self.rowsize, 23, self.xl_Text1[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 23, 'Empty', self.__style9)
                    if self.xl_Text2[Iteration_Count]:
                        self.ws.write(self.rowsize, 24, self.xl_Text2[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 24, 'Empty', self.__style9)
                    if self.xl_Text10[Iteration_Count]:
                        self.ws.write(self.rowsize, 25, self.xl_Text10[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 25, 'Empty', self.__style9)
                    if self.xl_Text11[Iteration_Count]:
                        self.ws.write(self.rowsize, 26, self.xl_Text11[Iteration_Count], self.__style9)
                    else:
                        self.ws.write(self.rowsize, 26, 'Empty', self.__style9)
                    # self.ws.write(self.rowsize, 26, self.xl_Text12[Iteration_Count], self.__style1)

                    # ----------------------
                    #   Write Output Data
                    # ----------------------
                    self.rowsize += 1  # Row increment
                    self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                    isSuccess = True
                    if candidate_id and self.xl_PassportNo == str(self.api_candidate_PassportNo)\
                        and self.xl_PanNo == str(self.api_candidate_PanNo)\
                        and self.xl_Gender == self.api_candidate_Gender:
                        self.ws.write(self.rowsize, 1, 'Pass', self.__style9)
                    else:
                        self.ws.write(self.rowsize, 1, 'Fail', self.__style3)
                        isSuccess = False
                    # self.ws.write(self.rowsize, 2, self.ID)

                    # ------------------------------------------------------------------
                    # Comparing API Data with Excel Data and Printing into Output Excel
                    # ------------------------------------------------------------------

                    if self.api_candidate_CandidateName:
                        if self.xl_Name[Iteration_Count] == str(self.api_candidate_CandidateName):
                            self.ws.write(self.rowsize, 4, str(self.api_candidate_CandidateName), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 4, str(self.api_candidate_CandidateName), self.__style7)
                    else:
                        self.ws.write(self.rowsize, 4, 'Empty', self.__style9)
                    # --------------------------------------------------------------------------
                    #   Email
                    # --------------------------------------------------------------------------
                    if self.api_candidate_Email:
                        if self.xl_Email[Iteration_Count] == str(self.api_candidate_Email):
                            self.ws.write(self.rowsize, 5, str(self.api_candidate_Email), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 5, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 5, 'Empty', self.__style9)
                    # -----------------------------------------------------------------------------
                    #   Mobile1
                    # -----------------------------------------------------------------------------

                    if self.api_candidate_Mobile1:
                        if str(self.xl_Mobile1[Iteration_Count]) == str(self.api_candidate_Mobile1):
                            self.ws.write(self.rowsize, 6, str(self.api_candidate_Mobile1), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 6, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 6, 'Empty', self.__style9)

                    if self.api_candidate_Gender:
                        if self.api_candidate_Gender == self.api_candidate_Gender:
                            self.ws.write(self.rowsize, 7, self.api_candidate_Gender, self.__style10)
                        else:
                            self.ws.write(self.rowsize, 7, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 7, 'Empty', self.__style9)

                    if self.api_candidate_DateOfBirth:
                        if str(self.xl_DateOfBirth[Iteration_Count]) == str(self.api_candidate_DateOfBirth):
                            self.ws.write(self.rowsize, 8, str(self.api_candidate_DateOfBirth), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 8, 'Fail', self.__style7)
                            isSuccess = False

                    else:
                        self.ws.write(self.rowsize, 8, 'Empty', self.__style9)

                    if self.api_candidate_BusinessUnit:
                        if self.xl_BusinessUnit[Iteration_Count] == self.api_candidate_BusinessUnit:
                            self.ws.write(self.rowsize, 9, self.api_candidate_BusinessUnit, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 9, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 9, 'Empty', self.__style9)

                    if self.api_candidate_Department:
                        if self.xl_Department[Iteration_Count] == self.api_candidate_Department:
                            self.ws.write(self.rowsize, 10, self.api_candidate_Department, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 10, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 10, 'Empty', self.__style9)

                    if self.api_candidate_OfferedRole:
                        if self.xl_OfferedRole[Iteration_Count] == self.api_candidate_OfferedRole:
                            self.ws.write(self.rowsize, 11, self.api_candidate_OfferedRole, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 11, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 11, 'Empty', self.__style9)

                    if self.api_candidate_PassportNo:
                        if self.api_candidate_PassportNo == str(self.api_candidate_PassportNo):
                            self.ws.write(self.rowsize, 12, str(self.api_candidate_PassportNo), self.__style10)
                        else:
                            self.ws.write(self.rowsize, 12, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 12, 'Empty', self.__style9)

                    if self.api_candidate_Level:
                        if self.xl_Level[Iteration_Count] == self.api_candidate_Level:
                            self.ws.write(self.rowsize, 13, self.api_candidate_Level, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 13, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 13, 'Empty', self.__style9)

                    if self.api_candidate_PanNo:
                        if self.api_candidate_PanNo == str(self.api_candidate_PanNo):
                            self.ws.write(self.rowsize, 14, str(self.api_candidate_PanNo), self.__style10)
                        else:
                            self.ws.write(self.rowsize, 14, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 14, 'Empty', self.__style9)

                    if self.api_candidate_OfferedLocation:
                        if self.xl_OfferedLocation[Iteration_Count] == self.api_candidate_OfferedLocation:
                            self.ws.write(self.rowsize, 15, self.api_candidate_OfferedLocation, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 15, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 15, 'Empty', self.__style9)

                    if self.api_candidate_ExpectedDOJ:
                        if self.xl_ExpectedDOJ[Iteration_Count] == str(self.api_candidate_ExpectedDOJ):
                            self.ws.write(self.rowsize, 16, str(self.api_candidate_ExpectedDOJ), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 16, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 16, 'Empty', self.__style9)

                    if self.api_candidate_TentativeDOJ:
                        if self.xl_TentativeDOJ[Iteration_Count] == str(self.api_candidate_TentativeDOJ):
                            self.ws.write(self.rowsize, 17, str(self.api_candidate_TentativeDOJ), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 17, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 17, 'Empty', self.__style9)

                    if self.api_candidate_Integer1:
                        if self.xl_Integer1[Iteration_Count] == self.api_candidate_Integer1:
                            self.ws.write(self.rowsize, 18, self.api_candidate_Integer1, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 18, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 18, 'Empty', self.__style9)

                    if self.api_candidate_Integer2:
                        if self.xl_Integer2[Iteration_Count] == self.api_candidate_Integer2:
                            self.ws.write(self.rowsize, 19, self.api_candidate_Integer2, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 19, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 19, 'Empty', self.__style9)

                    if self.api_candidate_Integer3:
                        if self.xl_Integer3[Iteration_Count] == self.api_candidate_Integer3:
                            self.ws.write(self.rowsize, 20, self.api_candidate_Integer3, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 20, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 20, 'Empty', self.__style9)

                    if self.api_candidate_Integer10:
                        if self.xl_Integer10[Iteration_Count] == self.api_candidate_Integer10:
                            self.ws.write(self.rowsize, 21, self.api_candidate_Integer10, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 21, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 21, 'Empty', self.__style9)

                    if self.api_candidate_Integer12:
                        if self.xl_Integer12[Iteration_Count] == self.api_candidate_Integer12:
                            self.ws.write(self.rowsize, 22, self.api_candidate_Integer12, self.__style9)
                        else:
                            self.ws.write(self.rowsize, 22, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 22, 'Empty', self.__style9)

                    if self.api_candidate_Text1:
                        if self.xl_Text1[Iteration_Count] == str(self.api_candidate_Text1):
                            self.ws.write(self.rowsize, 23, str(self.api_candidate_Text1), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 23, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 23, 'Empty', self.__style9)

                    if self.api_candidate_Text2:
                        if self.xl_Text2[Iteration_Count] == str(self.api_candidate_Text2):
                            self.ws.write(self.rowsize, 24, str(self.api_candidate_Text2), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 24, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 24, 'Empty', self.__style9)

                    if self.api_candidate_Text10:
                        if self.xl_Text10[Iteration_Count] == str(self.api_candidate_Text10):
                            self.ws.write(self.rowsize, 25, str(self.api_candidate_Text10), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 25, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 25, 'Empty', self.__style9)

                    if self.api_candidate_Text11:
                        if self.xl_Text11[Iteration_Count] == str(self.api_candidate_Text11):
                            self.ws.write(self.rowsize, 26, str(self.api_candidate_Text11), self.__style9)
                        else:
                            self.ws.write(self.rowsize, 26, 'Fail', self.__style7)
                            isSuccess = False
                    else:
                        self.ws.write(self.rowsize, 26, 'Empty', self.__style9)

                    # if self.api_candidate_Text12:
                    #     if self.xl_Text12[Iteration_Count] == str(self.api_candidate_Text12):
                    #         self.ws.write(self.rowsize, 26, str(self.api_candidate_Text12), self.__style9)
                    #     else:
                    #         self.ws.write(self.rowsize, 26, 'Fail', self.__style7)
                    # else:
                    #     self.ws.write(self.rowsize, 26, 'Api Response is null', self.__style7)
                    #
                    # if isSuccess == False:
                    #     self.ws.write(self.rowsize, 27, 'Fail', self.__style7)

                    self.rowsize += 1  # Row increment
                    ob.wb_Result.save(
                        '/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                        '/BulkUpdate_kmbdemo_output.xls')

        except Exception as e:
            print(e)
            print("DB Connection Error - Exception Block")

    def over_status(self):
        self.ws.write(0, 0, 'BU_Candidates', self.__style9)
        if self.Expected_success_cases == self.Expected_success_cases:
            self.ws.write(0, 1, 'Fail', self.__style9)
        else:
            self.ws.write(0, 1, 'pass', self.__style7)

        self.ws.write(0, 3, 'StartTime', self.__style9)
        self.ws.write(0, 4, self.start_time)
        ob.wb_Result.save('/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                          '/BulkUpdate_kmbdemo_output.xls')


ob = UpdateCandidate()
ob.excel_data()

total_len = len(ob.xl_CandidateId)
for Iteration_Count in range(0, total_len):
    print("Iteration Count:- %s" % Iteration_Count)
    ob.api_call_bulk_update(Iteration_Count)
    # ob.Db_connection(Iteration_Count)
    # ob.output_excel(Iteration_Count)
ob.over_status()
