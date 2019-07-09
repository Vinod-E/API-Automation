import pymysql
import requests
import json
import xlrd
import datetime
import xlwt
import time


class Submit_Task:
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
        self.xl_FirstName = []
        self.xl_LastName = []
        self.xl_DateTime = []
        self.xl_DateTime01 = []
        self.xl_Colleges = []
        self.xl_Degree = []
        self.xl_FieldYouWant = []
        self.xl_Gender = []
        self.xl_ExtraCourse = []
        self.xl_OtherOptionalField = []
        self.xl_AddressPermanent = []
        self.xl_AddressCommunication = []
        self.xl_ProfilePIC = []
        self.xl_Documents = []
        self.xl_Date = []
        self.xl_Date01 = []
        self.xl_CurrentTime = []
        self.xl_SubmitTime = []
        self.xl_FirstNameM = []
        self.xl_LastNameM = []
        self.xl_DateTimeM = []
        self.xl_DateTime01M = []
        self.xl_CollegesM = []
        self.xl_DegreeM = []
        self.xl_FieldYouWantM = []
        self.xl_GenderM = []
        self.xl_ExtraCourseM = []
        self.xl_OtherOptionalFieldM = []
        self.xl_AddressPermanentM = []
        self.xl_AddressCommunicationM = []
        self.xl_DateM = []
        self.xl_Date01M = []
        self.xl_CurrentTimeM = []
        self.xl_SubmitTimeM = []
        self.xl_Expected_msg = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 37)))
    #
    #     # -------------------------------------------------------
    #     # Styles for Excel sheet Row, Column, Text - color, Font
    #     # -------------------------------------------------------
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
    #
    #     # -------------------------------------
    #     # Excel sheet write for Output results
    #     # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Submit_Task')
        self.rowsize = 2
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparison', 'Remark', 'CandidateID', 'Current_status',
                        'FirstName', 'LastName', 'DateTime', 'DateTime01', 'colleges', 'Degree', 'FieldYouWant',
                        'Gender',
                        'ExtraCourse', 'OtherOptionalField', 'AddressPermanent', 'AddressCommunication', 'ProfilePIC',
                        'Documents',
                        'Date', 'Date01', 'CurrentTime', 'SubmitTime', 'FirstNameM', 'LastNameM', 'DateTimeM',
                        'DateTime01M',
                        'CollegesM', 'DegreeM', 'FieldYouWantM', 'GenderM', 'ExtraCourseM', 'OtherOptionalFieldM',
                        'AddressPermanentM', 'AddressCommunicationM',
                        'DateM', 'Date01M', 'CurrentTimeM', 'SubmitTimeM']

        for headers in excelheaders:
            if headers in ['Comparison', 'Remark', 'Current_status', 'CandidateID']:
                self.ws.write(1, index, headers, self.__style2)
            else:
                self.ws.write(1, index, headers, self.__style0)
            index += 1

    def excel_data(self):
    # ----------------
    # Excel Data Read
    # ----------------

         self.wb = xlrd.open_workbook('/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/Upload_candidates'
                                      '/Submit_Task_03.xls')
         self.sheet1 = self.wb.sheet_by_index(0)
         for i in range(1, self.sheet1.nrows):
             number = i
             rows = self.sheet1.row_values(number)
            # ----------------------------------------------------------
            # Task Details for submit
            # ----------------------------------------------------------

             if not rows[0]:
                self.xl_FirstName.append(None)
             else:
                self.xl_FirstName.append(str(rows[0]))

             if not rows[1]:
                self.xl_LastName.append(None)
             else:
                self.xl_LastName.append(str(rows[1]))

             if not rows[2]:
                self.xl_DateTime.append(None)
             else:
                # self.xl_DateTime.append(rows[2])
                DateTime = self.sheet1.cell_value(rowx=(i), colx=2)
                self.DateTime = datetime.datetime(*xlrd.xldate_as_tuple(DateTime, self.wb.datemode))
                self.DateTime = self.DateTime.strftime("%d-%b-%Y %H:%M")
                self.xl_DateTime.append(self.DateTime)
                print(self.DateTime)

             if not rows[3]:
                self.xl_DateTime01.append(None)
             else:
                # self.xl_DateTime01.append(str(rows[3]))
                DateTime01 = self.sheet1.cell_value(rowx=(i), colx=3)
                self.DateTime01 = datetime.datetime(*xlrd.xldate_as_tuple(DateTime01, self.wb.datemode))
                self.DateTime01 = self.DateTime01.strftime("%d-%b-%Y %H:%M")
                self.xl_DateTime01.append(self.DateTime01)

             if not rows[4]:
                self.xl_Colleges.append(None)
             else:
                self.xl_Colleges.append(int(rows[4]))

             if not rows[5]:
                self.xl_Degree.append(None)

             else:
                self.xl_Degree.append(int(rows[5]))
             if not rows[6]:
                self.xl_FieldYouWant.append(None)
             else:
                self.xl_FieldYouWant.append(rows[6])

             if not rows[7]:
                self.xl_Gender.append(None)
             else:
                self.xl_Gender.append(rows[7])

             if not rows[8]:
                self.xl_ExtraCourse.append(None)
             else:
                self.xl_ExtraCourse.append(rows[8])

             if not rows[9]:
                self.xl_OtherOptionalField.append(None)
             else:
                self.xl_OtherOptionalField.append(rows[9])

             if not rows[10]:
                self.xl_AddressPermanent.append(None)
             else:
                self.xl_AddressPermanent.append(rows[10])

             if not rows[11]:
                self.xl_AddressCommunication.append(None)
             else:
                self.xl_AddressCommunication.append(rows[11])

             if not rows[12]:
                self.xl_ProfilePIC.append(None)
             else:
                self.xl_ProfilePIC.append(rows[12])

             if not rows[13]:
                self.xl_Documents.append(None)
             else:
                self.xl_Documents.append(rows[13])

             if not rows[14]:
                self.xl_Date.append(None)

             else:
                Date = self.sheet1.cell_value(rowx=(i), colx=14)
                self.Date = datetime.datetime(*xlrd.xldate_as_tuple(Date, self.wb.datemode))
                self.Date = self.Date.strftime("%d-%b-%Y %H:%M").replace('00:00', "")
                self.xl_Date.append(self.Date)
                print(self.xl_Date)

             if not rows[15]:
                self.xl_Date01.append(None)

             else:
                Date01 = self.sheet1.cell_value(rowx=(i), colx=15)
                self.Date01 = datetime.datetime(*xlrd.xldate_as_tuple(Date01, self.wb.datemode))
                self.Date01 = self.Date01.strftime("%d-%b-%Y %H:%M").replace('00:00', "")
                self.xl_Date01.append(self.Date01)
                print(self.xl_Date01)

             if not rows[16]:
                self.xl_CurrentTime.append(None)
             else:
                CurrentTime = self.sheet1.cell_value(rowx=(i), colx=16)
                self.CurrentTime = datetime.datetime(*xlrd.xldate_as_tuple(CurrentTime, self.wb.datemode))
                self.CurrentTime = self.CurrentTime.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                self.xl_CurrentTime.append(self.CurrentTime)
                print(self.CurrentTime)

             if not rows[17]:
                self.xl_SubmitTime.append(None)
             else:
                SubmitTime = self.sheet1.cell_value(rowx=(i), colx=17)
                self.SubmitTime = datetime.datetime(*xlrd.xldate_as_tuple(SubmitTime, self.wb.datemode))
                self.SubmitTime = self.SubmitTime.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                self.xl_SubmitTime.append(self.SubmitTime)
                print(self.SubmitTime)

             if not rows[18]:
                self.xl_FirstNameM.append(None)
             else:
                self.xl_FirstNameM.append(rows[18])

             if not rows[19]:
                self.xl_LastNameM.append(None)
             else:
                self.xl_LastNameM.append(rows[19])

             if not rows[20]:
                self.xl_DateTimeM.append(None)
             else:
                DateTimeM = self.sheet1.cell_value(rowx=(i), colx=20)
                self.DateTimeM = datetime.datetime(*xlrd.xldate_as_tuple(DateTimeM, self.wb.datemode))
                self.DateTimeM = self.DateTimeM.strftime("%d-%b-%Y %H:%M").replace('%H:%M', "")
                self.xl_DateTimeM.append(self.DateTimeM)

             if not rows[21]:
                self.xl_Date01M.append(None)
             else:
                DateTime01M = self.sheet1.cell_value(rowx=(i), colx=21)
                self.DateTime01M = datetime.datetime(*xlrd.xldate_as_tuple(DateTime01M, self.wb.datemode))
                self.DateTime01M = self.DateTime01M.strftime("%d-%b-%Y %H:%M").replace('%H:%M', "")
                self.xl_DateTime01M.append(self.DateTime01M)

             if not rows[22]:
                self.xl_CollegesM.append(None)
             else:
                self.xl_CollegesM.append(int(rows[22]))

             if not rows[23]:
                self.xl_DegreeM.append(None)
             else:
                self.xl_DegreeM.append(int(rows[23]))

             if not rows[24]:
                self.xl_FieldYouWantM.append(None)
             else:
                self.xl_FieldYouWantM.append(rows[24])

             if not rows[25]:
                self.xl_GenderM.append(None)
             else:
                self.xl_GenderM.append(rows[25])

             if not rows[26]:
                 self.xl_ExtraCourseM.append(None)
             else:
                 self.xl_ExtraCourseM.append(rows[26])

             if not rows[27]:
                 self.xl_OtherOptionalFieldM.append(None)
             else:
                 self.xl_OtherOptionalFieldM.append(rows[27])

             if not rows[28]:
                self.xl_AddressPermanentM.append(None)
             else:
                self.xl_AddressPermanentM.append(rows[28])
             if not rows[29]:
                self.xl_AddressCommunicationM.append(None)
             else:
                self.xl_AddressCommunicationM.append(rows[29])

             if not rows[30]:
                self.xl_DateM.append(None)
             else:
                DateM = self.sheet1.cell_value(rowx=(i), colx=30)
                self.DateM = datetime.datetime(*xlrd.xldate_as_tuple(DateM, self.wb.datemode))
                self.DateM = self.DateM.strftime("%d-%m-%Y")
                self.xl_DateM.append(self.Date)
                print(self.xl_DateM)

             if not rows[31]:
                self.xl_Date01M.append(None)
             else:
                Date01M = self.sheet1.cell_value(rowx=(i), colx=31)
                self.Date01M = datetime.datetime(*xlrd.xldate_as_tuple(Date01M, self.wb.datemode))
                self.Date01M = self.Date01M.strftime("%d-%m-%Y")
                self.xl_Date01M.append(self.Date)
                print(self.xl_Date01M)

             if not rows[32]:
                self.xl_CurrentTimeM.append(None)
             else:
                CurrentTimeM = self.sheet1.cell_value(rowx=(i), colx=32)
                self.CurrentTimeM = datetime.datetime(*xlrd.xldate_as_tuple(CurrentTimeM, self.wb.datemode))
                self.CurrentTimeM = self.CurrentTimeM.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                self.xl_CurrentTimeM.append(self.CurrentTimeM)
                print(self.CurrentTimeM)

             if not rows[33]:
                self.xl_SubmitTime.append(None)
             else:
                SubmitTimeM = self.sheet1.cell_value(rowx=(i), colx=33)
                self.SubmitTimeM = datetime.datetime(*xlrd.xldate_as_tuple(SubmitTimeM, self.wb.datemode))
                self.SubmitTimeM = self.SubmitTimeM.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                self.xl_SubmitTimeM.append(self.SubmitTimeM)
                print(self.SubmitTimeM)

                # self.xl_Address[0], self.xl_Resume[0], self.Date.replace('00:00', ""), self.CurrentTime]

    def submit_task_by_candidate(self, total_len):
        try:

            conn = pymysql.connect(host='35.154.36.218',
                                           database='appserver_core',
                                           user='qauser',
                                           password='qauser')
            cur1 = conn.cursor()
            self.total = 1
            self.cand_list = []
            print(cur1)
            b = "select candidate_id from assigned_tasks where task_id =1357"
            print(b)
            cur1.execute(b)
            candidateids = cur1.fetchall()
            self.control_values = []
            m = 0
            for row in range(0, total_len):
                zm = [self.xl_FirstName[row], self.xl_LastName[row], self.DateTime, self.DateTime01,
                      self.xl_Colleges[row], self.xl_Degree[row], self.xl_FieldYouWant[row], self.xl_Gender[row],
                      self.xl_ExtraCourse[row], self.xl_OtherOptionalField[row], self.xl_AddressPermanent[row],
                      self.xl_AddressCommunication[row], self.xl_ProfilePIC[row], self.xl_Documents[row],
                      self.Date.replace('00:00', ""), self.Date01.replace('00:00', ""), self.CurrentTime,
                      self.SubmitTime, self.xl_FirstNameM[row], self.xl_LastNameM[row], self.DateTimeM,
                      self.DateTime01M, self.xl_CollegesM[row], self.xl_DegreeM[row], self.xl_FieldYouWantM[row],
                      self.xl_GenderM[row], self.xl_ExtraCourseM[row], self.xl_OtherOptionalFieldM[row],
                      self.xl_AddressPermanentM[row], self.xl_AddressCommunicationM[row],
                      self.DateM.replace('00:00', ""),
                      self.Date01M.replace('00:00', ""),
                      self.CurrentTimeM, self.SubmitTimeM]
                self.control_values.append(zm)
            for x in candidateids:
                self.cand_list.append(x[0])

            self.header = {"Content-type": "application/json"}
            self.data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "staffingautomation", "UserName": "admin"}
            response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=self.header,
                                     data=json.dumps(self.data), verify=False)
            self.Token = response.json().get("Token")

            self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.Token}
            for y in self.cand_list:
                self.get_by_id = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False,
                                                     "MaxResults": 500, "PageNo": 1}, "CandidateId": y,
                                  "UserDepartmentId": 0}

                r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",
                                  headers=self.headers, data=json.dumps(self.get_by_id, default=str), verify=False)

                resp_dict = json.loads(r.content)
                # self.CandidateId = resp_dict.get('CandidateId')
                # print self.CandidateId

                time.sleep(1)
                self.status = resp_dict['status']
                Candidate_Task_Collection = resp_dict['CandidateTaskCollection']

                print(Candidate_Task_Collection)
                for a in Candidate_Task_Collection:
                    if a.get('TaskId') == 1357 and a['Status'] == 0:
                        form_id = a['FormId']
                        task_id = a['TaskId']
                        self.ticket = a['TicketNumber']
                        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.Token}
                        self.data = {"FormId": a['FormId'], "FilledFormId": None, "IsControlListsRequired": True,
                                     "CandidateId": y, "TaskId": a['TaskId']}

                        r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-form-by-id/",
                                          headers=self.headers,
                                          data=json.dumps(self.data, default=str), verify=False)
                        print(self.data)
                        resp_dict = json.loads(r.content)
                        self.status = resp_dict['status']
                        self.FormConfigurations = resp_dict['FormConfigurations']

                        datadict = []
                        k = 0
                        control_value = self.control_values[m]
                        for z in self.FormConfigurations:
                            datadict.append({"ControlId": z["Id"], "FormControlType": 0, "DisclaimerCheck": False,
                                             "ControlValue": control_value[k]})
                            k += 1
                        # print datadict
                        self.r = {"FormControlValues": datadict, "TicketId": self.ticket,
                                  "IsCoordinatorSubmittingBehalfOfCandidate": True}
                        print(self.data)

                        f = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/submit-form/", headers=self.headers,
                                          data=json.dumps(self.r, default=str), verify=False)
                        submit_candidate_response_resp_dict = json.loads(f.content)

                        print(submit_candidate_response_resp_dict)
                        time.sleep(1)
                        self.status = submit_candidate_response_resp_dict['status']
                        data = submit_candidate_response_resp_dict.get('ValidationMessage')
                        if data and len(data):
                            self.ws.write(self.rowsize, 1, data, self.__style3)
                            self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                            self.ws.write(self.rowsize, 3, self.xl_FirstName[m], self.__style1)
                            self.ws.write(self.rowsize, 4, self.xl_LastName[m], self.__style1)
                            self.ws.write(self.rowsize, 5, self.xl_DateTime[m], self.__style1)
                            self.ws.write(self.rowsize, 6, self.xl_DateTime01[m], self.__style1)
                            self.ws.write(self.rowsize, 7, self.xl_Colleges[m], self.__style1)
                            self.ws.write(self.rowsize, 8, self.xl_Degree[m], self.__style1)
                            self.ws.write(self.rowsize, 9, self.xl_FieldYouWant[m], self.__style1)
                            self.ws.write(self.rowsize, 10, self.xl_Gender[m], self.__style1)
                            self.ws.write(self.rowsize, 11, self.xl_ExtraCourse[m], self.__style1)
                            self.ws.write(self.rowsize, 12, self.xl_OtherOptionalField[m], self.__style1)
                            self.ws.write(self.rowsize, 13, self.xl_AddressPermanent[m], self.__style1)
                            self.ws.write(self.rowsize, 14, self.xl_AddressCommunication[m], self.__style1)
                            self.ws.write(self.rowsize, 15, self.xl_ProfilePIC[m], self.__style1)
                            self.ws.write(self.rowsize, 16, self.xl_Documents[m], self.__style1)
                            self.ws.write(self.rowsize, 17, self.xl_Date[m], self.__style1)
                            self.ws.write(self.rowsize, 18, self.xl_Date01[m], self.__style1)
                            self.ws.write(self.rowsize, 19, self.xl_CurrentTime[m], self.__style1)
                            self.ws.write(self.rowsize, 20, self.xl_SubmitTime[m], self.__style1)
                            self.ws.write(self.rowsize, 21, self.xl_FirstNameM[m], self.__style1)
                            self.ws.write(self.rowsize, 22, self.xl_LastNameM[m], self.__style1)
                            self.ws.write(self.rowsize, 23, self.xl_DateTimeM[m], self.__style1)
                            self.ws.write(self.rowsize, 24, self.DateTime01M[m], self.__style1)
                            self.ws.write(self.rowsize, 25, self.xl_CollegesM[m], self.__style1)
                            self.ws.write(self.rowsize, 26, self.xl_DegreeM[m], self.__style1)
                            self.ws.write(self.rowsize, 27, self.xl_FieldYouWantM[m], self.__style1)
                            self.ws.write(self.rowsize, 28, self.xl_GenderM[m], self.__style1)
                            self.ws.write(self.rowsize, 29, self.xl_ExtraCourseM[m], self.__style1)
                            self.ws.write(self.rowsize, 30, self.xl_OtherOptionalFieldM[m], self.__style1)
                            self.ws.write(self.rowsize, 31, self.xl_AddressPermanentM[m], self.__style1)
                            self.ws.write(self.rowsize, 32, self.xl_AddressCommunicationM[m], self.__style1)
                            self.ws.write(self.rowsize, 33, self.xl_DateM[m], self.__style1)
                            self.ws.write(self.rowsize, 34, self.xl_Date01M[m], self.__style1)
                            self.ws.write(self.rowsize, 35, self.xl_CurrentTimeM[m], self.__style1)
                            self.ws.write(self.rowsize, 36, self.xl_SubmitTimeM[m], self.__style1)
                            self.rowsize += 1  # Row increment
                            self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
                            self.rowsize += 1  # Row increment
                            ob.wb_Result.save(
                                '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                                '/Submit_Task.xls')


                        else:

                            r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",
                                              headers=self.headers,
                                              data=json.dumps(self.data, default=str), verify=False)
                            print("Below is Original Data")
                            print(self.data)
                            print("Below1 is Original Data")
                            resp_dict = json.loads(r.content)

                            print(resp_dict)

                            time.sleep(1)
                            self.status = resp_dict['status']
                            Candidate_Task_Collection01 = resp_dict['CandidateTaskCollection']
                            print(Candidate_Task_Collection01)

                            for d in Candidate_Task_Collection01:
                                if d['Status'] == 2:
                                    form_id = d['FormId']
                                    task_id = d['TaskId']
                                    formfilled = d['FilledFormId']
                                    CandidateId = d['CandidateId']
                                    # self.ticket =  d['TicketNumber']
                                    self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.Token}
                                    self.t = {"FormId": d['FormId'], "FilledFormId": d['FilledFormId'],
                                              "IsControlListsRequired": True, "TaskId": d['TaskId'], "CandidateId": y}
                                    r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-form-by-id/",
                                                      headers=self.headers,
                                                      data=json.dumps(self.t, default=str), verify=False)
                                    print(self.t)
                                    resp_dict = json.loads(r.content)
                                    FormConfigurations = resp_dict['FormConfigurations']
                                    print(FormConfigurations)
                                    self.a = []
                                    for g in FormConfigurations:
                                        api_id = g['Id']
                                        api_value = g['Value']
                                        self.c = self.a.append(api_value)
                                    self.api_FirstName = self.a[0]
                                    self.api_LastName = self.a[1]
                                    self.api_DateTime = self.a[2]
                                    self.api_DateTime01 = self.a[3]
                                    self.api_Colleges = int(self.a[4])
                                    self.api_Degree = int(self.a[5])
                                    self.api_FieldYouWant = self.a[6]
                                    self.api_Gender = self.a[7]
                                    self.api_ExtraCourse = self.a[8]
                                    self.api_OtherOptionalField = self.a[9]
                                    self.api_AddressPermanent = self.a[10]
                                    self.api_AddressCommunication = self.a[11]
                                    self.api_ProfilePIC = self.a[12]
                                    self.api_Documents = self.a[13]
                                    self.api_Date = self.a[14].replace('00:00', "")
                                    self.api_Date01 = self.a[15].replace('00:00', "")
                                    self.api_CurrentTime = self.a[16].replace('30-Dec-2018', "")
                                    self.api_SubmitTime = self.a[17].replace('30-Dec-2018', "")
                                    self.api_FirstNameM = self.a[18]
                                    self.api_LastNameM = self.a[19]
                                    self.api_DateTimeM = self.a[20]
                                    self.api_DateTime01M = self.a[21]
                                    self.api_CollegesM = int(self.a[22])
                                    self.api_DegreeM = int(self.a[23])
                                    self.api_FieldYouWantM = self.a[24]
                                    self.api_GenderM = self.a[25]
                                    self.api_ExtraCourseM = self.a[26]
                                    self.api_OtherOptionalFieldM = self.a[27]
                                    self.api_AddressPermanentM = self.a[28]
                                    self.api_AddressCommunicationM = self.a[29]
                                    self.api_DateM = self.a[30].replace('00:00', "")
                                    self.api_Date01M = self.a[31].replace('00:00', "")
                                    self.api_CurrentTimeM = self.a[32].replace('30-Dec-2018', "")
                                    self.api_SubmitTimeM = self.a[33].replace('30-Dec-2018', "")

                                    self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
                                    self.ws.write(self.rowsize, 1, 'Excel Input', self.__style9)
                                    if self.xl_FirstName[m]:
                                        self.ws.write(self.rowsize, 4, self.xl_FirstName[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 4, 'Empty', self.__style9)
                                    if self.xl_LastName[m]:
                                        self.ws.write(self.rowsize, 5, self.xl_LastName[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 5, 'Empty', self.__style9)
                                    if self.xl_DateTime[m]:
                                        self.ws.write(self.rowsize, 6, self.xl_DateTime[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 6, 'Empty', self.__style9)
                                    if self.xl_DateTime01[m]:
                                        self.ws.write(self.rowsize, 7, self.xl_DateTime01[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 7, 'Empty', self.__style9)
                                    if self.xl_Colleges[m]:
                                        self.ws.write(self.rowsize, 8, self.xl_Colleges[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 8, 'Empty', self.__style9)
                                    if self.xl_Degree[m]:
                                        self.ws.write(self.rowsize, 9, self.xl_Degree[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 9, 'Empty', self.__style9)
                                    if self.xl_FieldYouWant[m]:
                                        self.ws.write(self.rowsize, 10, self.xl_FieldYouWant[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 10, 'Empty', self.__style9)
                                    if self.xl_Gender[m]:
                                        self.ws.write(self.rowsize, 11, self.xl_Gender[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 11, 'Empty', self.__style9)
                                    if self.xl_ExtraCourse[m]:
                                        self.ws.write(self.rowsize, 12, self.xl_ExtraCourse[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 12, 'Empty', self.__style9)
                                    if self.xl_OtherOptionalField[m]:
                                        self.ws.write(self.rowsize, 13, self.xl_OtherOptionalField[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 13, 'Empty', self.__style9)
                                    if self.xl_AddressPermanent[m]:
                                        self.ws.write(self.rowsize, 14, self.xl_AddressPermanent[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 14, 'Empty', self.__style9)
                                    if self.xl_AddressCommunication[m]:
                                        self.ws.write(self.rowsize, 15, self.xl_AddressCommunication[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 15, 'Empty', self.__style9)
                                    if self.xl_ProfilePIC[m]:
                                        self.ws.write(self.rowsize, 16, self.xl_ProfilePIC[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 16, 'Empty', self.__style9)
                                    if self.xl_Documents[m]:
                                        self.ws.write(self.rowsize, 17, self.xl_Documents[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 17, 'Empty', self.__style9)
                                    if  self.xl_Date[m]:
                                        self.ws.write(self.rowsize, 18, self.xl_Date[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 18, 'Empty', self.__style9)
                                    if self.xl_Date01[m]:
                                        self.ws.write(self.rowsize, 19, self.xl_Date01[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 19, 'Empty', self.__style9)
                                    if self.xl_CurrentTime[m]:
                                        self.ws.write(self.rowsize, 20, self.xl_CurrentTime[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 20, 'Empty', self.__style9)
                                    if self.xl_SubmitTime[m]:
                                        self.ws.write(self.rowsize, 21, self.xl_SubmitTime[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 21, 'Empty', self.__style9)
                                    if self.xl_FirstNameM[m]:
                                        self.ws.write(self.rowsize, 22, self.xl_FirstNameM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 22, 'Empty', self.__style9)
                                    if self.xl_LastNameM[m]:
                                        self.ws.write(self.rowsize, 23, self.xl_LastNameM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 23, 'Empty', self.__style9)
                                    if self.xl_DateTimeM[m]:
                                        self.ws.write(self.rowsize, 24, self.xl_DateTimeM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 24, 'Empty', self.__style9)
                                    if self.xl_DateTime01M[m]:
                                        self.ws.write(self.rowsize, 25, self.xl_DateTime01M[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 25, 'Empty', self.__style9)
                                    if self.xl_CollegesM[m]:
                                        self.ws.write(self.rowsize, 26, self.xl_CollegesM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 26, 'Empty', self.__style9)
                                    if self.xl_DegreeM[m]:
                                        self.ws.write(self.rowsize, 27, self.xl_DegreeM[m], self.__style1)
                                    else:
                                        self.ws.write(self.rowsize, 27, 'Empty', self.__style9)
                                    if self.xl_FieldYouWantM[m]:
                                        self.ws.write(self.rowsize, 28, self.xl_FieldYouWantM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 28, 'Empty', self.__style9)
                                    if self.xl_GenderM[m]:
                                        self.ws.write(self.rowsize, 29, self.xl_GenderM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 29, 'Empty', self.__style9)
                                    if self.xl_ExtraCourseM[m]:
                                        self.ws.write(self.rowsize, 30, self.xl_ExtraCourseM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 30, 'Empty', self.__style9)
                                    if self.xl_OtherOptionalFieldM[m]:
                                        self.ws.write(self.rowsize, 31, self.xl_OtherOptionalFieldM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 31, 'Empty', self.__style9)
                                    if self.xl_AddressPermanentM[m]:
                                        self.ws.write(self.rowsize, 32, self.xl_AddressPermanentM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 32, 'Empty', self.__style9)
                                    if self.xl_AddressCommunication[m]:
                                        self.ws.write(self.rowsize, 33, self.xl_AddressCommunication[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 33, 'Empty', self.__style9)
                                    if self.xl_DateM[m]:
                                        self.ws.write(self.rowsize, 34, self.xl_DateM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 34, 'Empty', self.__style9)
                                    if self.xl_Date01M[m]:
                                        self.ws.write(self.rowsize, 35, self.xl_Date01M[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 35, 'Empty', self.__style9)
                                    if self.xl_CurrentTimeM[m]:
                                        self.ws.write(self.rowsize, 36, self.xl_CurrentTimeM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 36, 'Empty', self.__style9)
                                    if self.xl_SubmitTimeM[m]:
                                        self.ws.write(self.rowsize, 37, self.xl_SubmitTimeM[m], self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 37, 'Empty', self.__style9)
                                    # ----------------------
                                    #   Write Output Data
                                    # ----------------------
                                    self.rowsize += 1  # Row increment
                                    self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

                                    if self.cand_list:
                                        self.ws.write(self.rowsize, 1, 'Submitted Successfully', self.__style9)
                                        self.ws.write(self.rowsize, 2, 'Pass', self.__style9)
                                    else:
                                        self.ws.write(self.rowsize, 1, 'Candidate Failed to Submit', self.__style3)
                                        self.ws.write(self.rowsize, 2, 'Fail', self.__style7)

                                    self.ws.write(self.rowsize, 3, CandidateId)

                                    # ------------------------------------------------------------------
                                    # Comparing API Data with Excel Data and Printing into Output Excel
                                    # ------------------------------------------------------------------

                                    if self.api_FirstName:
                                        if str(self.xl_FirstName[m]) == str(self.api_FirstName):
                                            self.ws.write(self.rowsize, 4, str(self.api_FirstName), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 4, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 4, 'Empty', self.__style9)
                                    # __________________________________________________________________________
                                    # LastName
                                    # __________________________________________________________________________

                                    if self.api_LastName:
                                        if self.xl_LastName[m] == str(self.api_LastName):
                                            self.ws.write(self.rowsize, 5, str(self.api_LastName), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 5, str(self.api_LastName), self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 5, 'Empty', self.__style9)
                                    # --------------------------------------------------------------------------
                                    #   DateTime
                                    # --------------------------------------------------------------------------
                                    if self.api_DateTime:
                                        if self.xl_DateTime[m] == str(self.api_DateTime):
                                            self.ws.write(self.rowsize, 6, str(self.api_DateTime), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 6, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 6, 'Empty', self.__style9)
                                    # -----------------------------------------------------------------------------
                                    #   DateTime01
                                    # -----------------------------------------------------------------------------

                                    if self.api_DateTime01:
                                        if str(self.xl_DateTime01[m]) == str(self.api_DateTime01):
                                            self.ws.write(self.rowsize, 7, str(self.api_DateTime01), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 7, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 7, 'Empty', self.__style9)

                                    # ------------------------------------------------------------------------------
                                    # Colleges
                                    # ------------------------------------------------------------------------------

                                    if self.api_Colleges:
                                        if self.xl_Colleges[m] == (self.api_Colleges):
                                            self.ws.write(self.rowsize, 8, (self.api_Colleges), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 8, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 8, 'Empty', self.__style9)

                                    if self.api_Degree:
                                        if str(self.xl_Degree[m]) == str(self.api_Degree):
                                            self.ws.write(self.rowsize, 9, str(self.api_Degree), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 9, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 9, 'Empty', self.__style9)

                                    if self.api_FieldYouWant:
                                        if self.xl_FieldYouWant[m] == self.api_FieldYouWant:
                                            self.ws.write(self.rowsize, 10, self.api_FieldYouWant, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 10, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 10, 'Empty', self.__style9)

                                    if self.api_Gender:
                                        if self.xl_Gender[m] == self.api_Gender:
                                            self.ws.write(self.rowsize, 11, self.api_Gender, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 11, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 11, 'Empty', self.__style9)

                                    if self.api_ExtraCourse:
                                        if self.xl_ExtraCourse[m] == self.api_ExtraCourse:
                                            self.ws.write(self.rowsize, 12, self.api_ExtraCourse, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 12, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 12, 'Empty', self.__style9)

                                    if self.api_OtherOptionalField:
                                        if self.xl_OtherOptionalField[m] == self.api_OtherOptionalField:
                                            self.ws.write(self.rowsize, 13, self.api_OtherOptionalField, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 13, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 13, 'Empty', self.__style9)

                                    if self.api_AddressPermanent:
                                        if self.xl_AddressPermanent[m] == str(self.api_AddressPermanent):
                                            self.ws.write(self.rowsize, 14, str(self.api_AddressPermanent),
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 14, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 14, 'Empty', self.__style9)

                                    if self.api_AddressCommunication:
                                        if self.xl_AddressCommunication[m] == self.api_AddressCommunication:
                                            self.ws.write(self.rowsize, 15, self.api_AddressCommunication,
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 15, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 15, 'Empty', self.__style9)

                                    if self.api_ProfilePIC:
                                        if self.xl_ProfilePIC[m] == str(self.api_ProfilePIC):
                                            self.ws.write(self.rowsize, 16, str(self.api_ProfilePIC), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 16, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 16, 'Empty', self.__style9)

                                    if self.api_Documents:
                                        if self.xl_Documents[m] == self.api_Documents:
                                            self.ws.write(self.rowsize, 17, self.api_Documents, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 17, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 17, 'Empty', self.__style9)

                                    if self.api_Date:
                                        if self.xl_Date[m] == str(self.api_Date):
                                            self.ws.write(self.rowsize, 18, str(self.api_Date), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 18, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 18, 'Empty', self.__style9)

                                    if self.api_Date01:
                                        if self.xl_Date01[m] == str(self.api_Date01):
                                            self.ws.write(self.rowsize, 19, str(self.api_Date01), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 19, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 19, 'Empty', self.__style9)

                                    if self.api_CurrentTime:
                                        if self.xl_CurrentTime[m] == self.api_CurrentTime:
                                            self.ws.write(self.rowsize, 20, self.api_CurrentTime, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 20, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 20, 'Empty', self.__style9)

                                    if self.api_SubmitTime:
                                        if self.xl_SubmitTime[m] == self.api_SubmitTime:
                                            self.ws.write(self.rowsize, 21, self.api_SubmitTime, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 21, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 21, 'Empty', self.__style9)

                                    if self.api_FirstNameM:
                                        if self.xl_FirstNameM[m] == self.api_FirstNameM:
                                            self.ws.write(self.rowsize, 22, self.api_FirstNameM, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 22, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 22, 'Empty', self.__style9)

                                    if self.api_LastNameM:
                                        if self.xl_LastNameM[m] == self.api_LastNameM:
                                            self.ws.write(self.rowsize, 23, self.api_LastNameM, self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 23, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 23, 'Empty', self.__style9)

                                    if self.api_DateTimeM:
                                        if self.xl_DateTimeM[m] == (self.api_DateTimeM):
                                            self.ws.write(self.rowsize, 24, (self.api_DateTimeM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 24, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 24, 'Empty', self.__style9)

                                    if self.api_DateTime01M:
                                        if self.xl_DateTime01M[m] == (self.api_DateTime01M):
                                            self.ws.write(self.rowsize, 25, (self.api_DateTime01M), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 25, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 25, 'Empty', self.__style9)

                                    if self.api_CollegesM:
                                        if self.xl_CollegesM[m] == (self.api_CollegesM):
                                            self.ws.write(self.rowsize, 26, (self.api_CollegesM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 26, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 26, 'Empty', self.__style9)

                                    if self.api_DegreeM:
                                        if self.xl_DegreeM[m] == (self.api_DegreeM):
                                            self.ws.write(self.rowsize, 27, (self.api_DegreeM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 27, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 27, 'Empty', self.__style9)

                                    if self.api_FieldYouWantM:
                                        if self.xl_FieldYouWantM[m] == str(self.api_FieldYouWantM):
                                            self.ws.write(self.rowsize, 28, str(self.api_FieldYouWantM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 28, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 28, 'Empty', self.__style9)

                                    if self.api_GenderM:
                                        if self.xl_GenderM[m] == str(self.api_GenderM):
                                            self.ws.write(self.rowsize, 29, str(self.api_GenderM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 29, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 29, 'Empty', self.__style9)

                                    if self.api_ExtraCourseM:
                                        if self.xl_ExtraCourseM[m] == str(self.api_ExtraCourseM):
                                            self.ws.write(self.rowsize, 30, str(self.api_ExtraCourseM),
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 30, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 30, 'Empty', self.__style9)

                                    if self.api_OtherOptionalFieldM:
                                        if self.xl_OtherOptionalFieldM[m] == str(self.api_OtherOptionalFieldM):
                                            self.ws.write(self.rowsize, 31, str(self.api_OtherOptionalFieldM),
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 31, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 31, 'Empty', self.__style9)

                                    if self.api_AddressPermanentM:
                                        if self.xl_AddressPermanentM[m] == str(self.api_AddressPermanentM):
                                            self.ws.write(self.rowsize, 32, str(self.api_AddressPermanentM),
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 32, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 32, 'Empty', self.__style9)

                                    if self.api_AddressCommunicationM:
                                        if self.xl_AddressCommunicationM[m] == str(self.api_AddressCommunicationM):
                                            self.ws.write(self.rowsize, 33, str(self.api_AddressCommunicationM),
                                                          self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 33, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 33, 'Empty', self.__style9)

                                    if self.api_DateM:
                                        if self.xl_DateM[m] == (self.api_DateM):
                                            self.ws.write(self.rowsize, 34, (self.api_DateM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 34, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 34, 'Empty', self.__style9)

                                    if self.api_Date01M:
                                        if self.xl_Date01M[m] == (self.api_Date01M):
                                            self.ws.write(self.rowsize, 35, (self.api_Date01M), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 35, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 35, 'Empty', self.__style9)

                                    if self.api_CurrentTimeM:
                                        if self.xl_CurrentTimeM[m] == (self.api_CurrentTimeM):
                                            self.ws.write(self.rowsize, 36, (self.api_CurrentTimeM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 36, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 36, 'Empty', self.__style9)
                                    if self.api_SubmitTimeM:
                                        if self.xl_SubmitTimeM[m] == (self.api_SubmitTimeM):
                                            self.ws.write(self.rowsize, 37, (self.api_SubmitTimeM), self.__style9)
                                        else:
                                            self.ws.write(self.rowsize, 37, 'Fail', self.__style7)
                                    else:
                                        self.ws.write(self.rowsize, 37, 'Empty', self.__style9)

                                    self.rowsize += 1  # Row increment
                                    ob.wb_Result.save(
                                        '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates'
                                        '/Submit_Task.xls')

                                    rr = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",
                                                       headers=self.headers,
                                                       data=json.dumps(self.data, default=str), verify=False)

                                    resp_dict01 = json.loads(rr.content)
                                    time.sleep(1)
                                    self.status = resp_dict01['status']

                                    self.ee = {"AssignedUserTaskIds": [self.ticket],
                                               "Comments": "please Fill Correct Data", "TaskStatus": 3}

                                    gg = requests.post(
                                        "https://amsin.hirepro.in/py/pofu/api/v1/update-assigned-task-status/",
                                        headers=self.headers,
                                        data=json.dumps(self.ee, default=str), verify=False)


                                    r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",
                                                      headers=self.headers,
                                                      data=json.dumps(self.get_by_id, default=str), verify=False)

                                    resp_dict01 = json.loads(r.content)
                                    Candidate_Task_Collection02 = resp_dict01['CandidateTaskCollection']
                                    print(Candidate_Task_Collection02)
                                    for dd in Candidate_Task_Collection02:
                                        if dd.get('TaskId') == 1357 and dd['Status'] == 3:
                                            ob1 = SubmitUtility()
                                            table1 = ob1.read_from_excel(
                                                '/home/vinodkumar/PycharmProjects/API_Automation/Input Data/Pofu/'
                                                'Upload_candidates/Submit_Task_04.xls')
                                            datadict1 = []
                                            k = 0
                                            control_value1 = table1[m]
                                            for z in self.FormConfigurations:
                                                datadict1.append({"ControlId": z["Id"], "FormControlType": 0,
                                                                 "DisclaimerCheck": False,
                                                                 "ControlValue": control_value1[k]})
                                                k += 1
                                            # print datadict
                                            self.r = {"FormControlValues": datadict1, "TicketId": self.ticket,
                                                      "IsCoordinatorSubmittingBehalfOfCandidate": True}
                                            print(self.data)

                                            f = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/submit-form/",
                                                              headers=self.headers,
                                                              data=json.dumps(self.r, default=str), verify=False)

                                            r = requests.post(
                                                "https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",
                                                headers=self.headers,
                                                data=json.dumps(self.get_by_id, default=str), verify=False)
                                            resp_dict = json.loads(r.content)

                                            time.sleep(1)
                                            self.status = resp_dict['status']
                                            Candidate_Task_Collection01 = resp_dict['CandidateTaskCollection']

                                            for d in Candidate_Task_Collection01:
                                                if d['Status'] == 2:
                                                    form_id = d['FormId']
                                                    task_id = d['TaskId']
                                                    formfilled = d['FilledFormId']
                                                    # self.ticket =  d['TicketNumber']
                                                    self.headers = {"content-type": "application/json",
                                                                    "X-AUTH-TOKEN": self.Token}
                                                    self.t = {"FormId": d['FormId'], "FilledFormId": d['FilledFormId'],
                                                              "IsControlListsRequired": True, "TaskId": d['TaskId'],
                                                              "CandidateId": y}
                                                    r = requests.post(
                                                        "https://amsin.hirepro.in/py/pofu/api/v1/get-form-by-id/",
                                                        headers=self.headers,
                                                        data=json.dumps(self.t, default=str), verify=False)
                                                    resp_dict = json.loads(r.content)
                                                    FormConfigurations = resp_dict['FormConfigurations']
                                                    print(FormConfigurations)
                                                    self.a = []
                                                    for g in FormConfigurations:
                                                        api_id = g['Id']
                                                        api_value = g['Value']
                                                        self.c = self.a.append(api_value)
                                                    self.api_FirstName = self.a[0]
                                                    self.api_LastName = self.a[1]
                                                    self.api_DateTime = self.a[2]
                                                    self.api_DateTime01 = self.a[3]
                                                    self.api_Colleges = int(self.a[4])
                                                    self.api_Degree = int(self.a[5])
                                                    self.api_FieldYouWant = self.a[6]
                                                    self.api_Gender = self.a[7]
                                                    self.api_ExtraCourse = self.a[8]
                                                    self.api_OtherOptionalField = self.a[9]
                                                    self.api_AddressPermanent = self.a[10]
                                                    self.api_AddressCommunication = self.a[11]
                                                    self.api_ProfilePIC = self.a[12]
                                                    self.api_Documents = self.a[13]
                                                    self.api_Date = self.a[14].replace('00:00', "")
                                                    self.api_Date01 = self.a[15].replace('00:00', "")
                                                    self.api_CurrentTime = self.a[16].replace('30-Dec-2018', "")
                                                    self.api_SubmitTime = self.a[17].replace('30-Dec-2018', "")
                                                    self.api_FirstNameM = self.a[18]
                                                    self.api_LastNameM = self.a[19]
                                                    self.api_DateTimeM = self.a[20]
                                                    self.api_DateTime01M = self.a[21]
                                                    self.api_CollegesM = int(self.a[22])
                                                    self.api_DegreeM = int(self.a[23])
                                                    self.api_FieldYouWantM = self.a[24]
                                                    self.api_GenderM = self.a[25]
                                                    self.api_ExtraCourseM = self.a[26]
                                                    self.api_OtherOptionalFieldM = self.a[27]
                                                    self.api_AddressPermanentM = self.a[28]
                                                    self.api_AddressCommunicationM = self.a[29]
                                                    self.api_DateM = self.a[30].replace('00:00', "")
                                                    self.api_Date01M = self.a[31].replace('00:00', "")
                                                    self.api_CurrentTimeM = self.a[32].replace('30-Dec-2018',
                                                                                               "")
                                                    self.api_SubmitTimeM = self.a[33].replace('30-Dec-2018', "")

                                                    self.ws.write(self.rowsize, self.col, 'Input',
                                                                  self.__style4)
                                                    self.ws.write(self.rowsize, 1, 'Excel Input', self.__style9)
                                                    if control_value1[0]:
                                                        self.ws.write(self.rowsize, 4, control_value1[0], self.__style1)
                                                    else:
                                                        self.ws.write(self.rowsize, 4, 'Empty', self.__style9)
                                                    if control_value1[1]:
                                                        self.ws.write(self.rowsize, 5, control_value1[1], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 5, 'Empty', self.__style9)
                                                    if control_value1[2]:
                                                        self.ws.write(self.rowsize, 6, control_value1[2], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 6, 'Empty', self.__style9)
                                                    if control_value1[3]:
                                                        self.ws.write(self.rowsize, 7, control_value1[3], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 7, 'Empty', self.__style9)
                                                    if control_value1[4]:
                                                        self.ws.write(self.rowsize, 8, control_value1[4], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 8, 'Empty', self.__style9)
                                                    if control_value1[5]:
                                                        self.ws.write(self.rowsize, 9, control_value1[5], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 9, 'Empty', self.__style9)
                                                    if control_value1[6]:
                                                        self.ws.write(self.rowsize, 10, control_value1[6],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 10, 'Empty', self.__style9)
                                                    if control_value1[7]:
                                                        self.ws.write(self.rowsize, 11, control_value1[7], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 11, 'Empty', self.__style9)
                                                    if control_value1[8]:
                                                        self.ws.write(self.rowsize, 12, control_value1[8],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 12, 'Empty', self.__style9)
                                                    if control_value1[9]:
                                                        self.ws.write(self.rowsize, 13, control_value1[9],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 13, 'Empty', self.__style9)
                                                    if control_value1[10]:
                                                        self.ws.write(self.rowsize, 14, control_value1[10],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 14, 'Empty', self.__style9)
                                                    if control_value1[11]:
                                                        self.ws.write(self.rowsize, 15, control_value1[11],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 15, 'Empty', self.__style9)
                                                    if control_value1[12]:
                                                        self.ws.write(self.rowsize, 16, control_value1[12],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 16, 'Empty', self.__style9)
                                                    if control_value1[13]:
                                                        self.ws.write(self.rowsize, 17, control_value1[13], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 17, 'Empty', self.__style9)
                                                    if control_value1[14]:
                                                        self.ws.write(self.rowsize, 18, control_value1[14], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 18, 'Empty', self.__style9)
                                                    if control_value1[15]:
                                                        self.ws.write(self.rowsize, 19, control_value1[15], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 19, 'Empty', self.__style9)
                                                    if control_value1[16]:
                                                        self.ws.write(self.rowsize, 20, control_value1[16],
                                                                 self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 20, 'Empty', self.__style9)
                                                    if control_value1[17]:
                                                        self.ws.write(self.rowsize, 21, control_value1[17],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 21, 'Empty', self.__style9)
                                                    if control_value1[18]:
                                                        self.ws.write(self.rowsize, 22, control_value1[18],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 22, 'Empty', self.__style9)
                                                    if control_value1[19]:
                                                        self.ws.write(self.rowsize, 23, control_value1[19], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 23, 'Empty', self.__style9)
                                                    if control_value1[20]:
                                                        self.ws.write(self.rowsize, 24, control_value1[20], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 24, 'Empty', self.__style9)
                                                    if control_value1[21]:
                                                        self.ws.write(self.rowsize, 25, control_value1[21],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 25, 'Empty', self.__style9)
                                                    if control_value1[22]:
                                                        self.ws.write(self.rowsize, 26, control_value1[22], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 26, 'Empty', self.__style9)
                                                    if control_value1[23]:
                                                        self.ws.write(self.rowsize, 27, control_value1[23], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 27, 'Empty', self.__style9)
                                                    if control_value1[24]:
                                                        self.ws.write(self.rowsize, 28, control_value1[24],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 28, 'Empty', self.__style9)
                                                    if control_value1[25]:
                                                        self.ws.write(self.rowsize, 29, control_value1[25], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 29, 'Empty', self.__style9)
                                                    if control_value1[26]:
                                                        self.ws.write(self.rowsize, 30, control_value1[26],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 30, 'Empty', self.__style9)
                                                    if control_value1[27]:
                                                        self.ws.write(self.rowsize, 31, control_value1[27],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 31, 'Empty', self.__style9)
                                                    if control_value1[28]:
                                                        self.ws.write(self.rowsize, 32, control_value1[28],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 32, 'Empty', self.__style9)
                                                    if control_value1[29]:
                                                        self.ws.write(self.rowsize, 33, control_value1[29],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 33, 'Empty', self.__style9)
                                                    if control_value1[30]:
                                                        self.ws.write(self.rowsize, 34, control_value1[30], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 34, 'Empty', self.__style9)
                                                    if control_value1[31]:
                                                        self.ws.write(self.rowsize, 35, control_value1[31], self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 35, 'Empty', self.__style9)
                                                    if control_value1[32]:
                                                        self.ws.write(self.rowsize, 36, control_value1[32],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 36, 'Empty', self.__style9)
                                                    if control_value1[33]:
                                                        self.ws.write(self.rowsize, 37, control_value1[33],
                                                                  self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 37, 'Empty', self.__style9)
                                                    # ----------------------
                                                    #   Write Output Data
                                                    # ----------------------
                                                    self.rowsize += 1  # Row increment
                                                    self.ws.write(self.rowsize, self.col, 'Output', self.__style5)

                                                    if self.cand_list:
                                                        self.ws.write(self.rowsize, 1, 'Updated Successfully',
                                                                      self.__style9)
                                                        self.ws.write(self.rowsize, 2, 'Pass', self.__style9)
                                                    else:
                                                        self.ws.write(self.rowsize, 1, 'Candidate Failed to Submit',
                                                                      self.__style3)
                                                    self.ws.write(self.rowsize, 2, 'Fail', self.__style7)

                                                    # ------------------------------------------------------------------
                                                    # Comparing API Data with Excel Data and Printing into Output Excel
                                                    # ------------------------------------------------------------------

                                                    if self.api_FirstName:
                                                        if str(control_value1[0]) == str(self.api_FirstName):
                                                            self.ws.write(self.rowsize, 4, str(self.api_FirstName),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 4, 'Fail', self.__style7)
                                                    else:

                                                        self.ws.write(self.rowsize, 4, 'Empty',
                                                                      self.__style9)
                                                    # __________________________________________________________________________
                                                    # LastName
                                                    # __________________________________________________________________________

                                                    if self.api_LastName:
                                                        if control_value1[1] == str(self.api_LastName):
                                                            self.ws.write(self.rowsize, 5, str(self.api_LastName),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 5, str(self.api_LastName),
                                                                          self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 5, 'Empty',
                                                                      self.__style9)
                                                    # --------------------------------------------------------------------------
                                                    #   DateTime
                                                    # --------------------------------------------------------------------------
                                                    if self.api_DateTime:
                                                        if control_value1[2] == str(self.api_DateTime):
                                                            self.ws.write(self.rowsize, 6, str(self.api_DateTime),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 6, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 6, 'Empty',
                                                                      self.__style9)
                                                    # -----------------------------------------------------------------------------
                                                    #   DateTime01
                                                    # -----------------------------------------------------------------------------

                                                    if self.api_DateTime01:
                                                        if str(control_value1[3]) == str(self.api_DateTime01):
                                                            self.ws.write(self.rowsize, 7, str(self.api_DateTime01),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 7, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 7, 'Empty',
                                                                      self.__style9)

                                                    # ------------------------------------------------------------------------------
                                                    # Colleges
                                                    # ------------------------------------------------------------------------------

                                                    if self.api_Colleges:
                                                        if control_value1[4] == (self.api_Colleges):
                                                            self.ws.write(self.rowsize, 8, (self.api_Colleges),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 8, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 8, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Degree:
                                                        if str(control_value1[5]) == str(self.api_Degree):
                                                            self.ws.write(self.rowsize, 9, str(self.api_Degree),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 9, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 9, 'Empty',
                                                                      self.__style9)

                                                    if self.api_FieldYouWant:
                                                        if control_value1[6] == self.api_FieldYouWant:
                                                            self.ws.write(self.rowsize, 10, self.api_FieldYouWant,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 10, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 10, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Gender:
                                                        if control_value1[7] == self.api_Gender:
                                                            self.ws.write(self.rowsize, 11, self.api_Gender,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 11, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 11, 'Empty',
                                                                      self.__style9)

                                                    if self.api_ExtraCourse:
                                                        if control_value1[8] == self.api_ExtraCourse:
                                                            self.ws.write(self.rowsize, 12, self.api_ExtraCourse,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 12, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 12, 'Empty',
                                                                      self.__style9)

                                                    if self.api_OtherOptionalField:
                                                        if control_value1[9] == self.api_OtherOptionalField:
                                                            self.ws.write(self.rowsize, 13, self.api_OtherOptionalField,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 13, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 13, 'Empty',
                                                                      self.__style9)

                                                    if self.api_AddressPermanent:
                                                        if control_value1[10] == str(
                                                                self.api_AddressPermanent):
                                                            self.ws.write(self.rowsize, 14,
                                                                          str(self.api_AddressPermanent), self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 14, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 14, 'Empty',
                                                                      self.__style9)

                                                    if self.api_AddressCommunication:
                                                        if control_value1[11] == self.api_AddressCommunication:
                                                            self.ws.write(self.rowsize, 15,
                                                                          self.api_AddressCommunication, self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 15, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 15, 'Empty',
                                                                      self.__style9)

                                                    if self.api_ProfilePIC:
                                                        if control_value1[12] == str(self.api_ProfilePIC):
                                                            self.ws.write(self.rowsize, 16, str(self.api_ProfilePIC),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 16, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 16, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Documents:
                                                        if control_value1[13] == self.api_Documents:
                                                            self.ws.write(self.rowsize, 17, self.api_Documents,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 17, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 17, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Date:
                                                        if control_value1[14] == str(self.api_Date):
                                                            self.ws.write(self.rowsize, 18, str(self.api_Date),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 18, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 18, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Date01:
                                                        if control_value1[15] == str(self.api_Date01):
                                                            self.ws.write(self.rowsize, 19, str(self.api_Date01),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 19, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 19, 'Empty',
                                                                      self.__style9)

                                                    if self.api_CurrentTime:
                                                        if control_value1[16] == self.api_CurrentTime:
                                                            self.ws.write(self.rowsize, 20, self.api_CurrentTime,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 20, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 20, 'Empty',
                                                                      self.__style9)

                                                    if self.api_SubmitTime:
                                                        if control_value1[17] == self.api_SubmitTime:
                                                            self.ws.write(self.rowsize, 21, self.api_SubmitTime,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 21, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 21, 'Empty',
                                                                      self.__style9)

                                                    if self.api_FirstNameM:
                                                        if control_value1[18] == self.api_FirstNameM:
                                                            self.ws.write(self.rowsize, 22, self.api_FirstNameM,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 22, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 22, 'Empty',
                                                                      self.__style9)

                                                    if self.api_LastNameM:
                                                        if control_value1[19] == self.api_LastNameM:
                                                            self.ws.write(self.rowsize, 23, self.api_LastNameM,
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 23, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 23, 'Empty',
                                                                      self.__style9)

                                                    if self.api_DateTimeM:
                                                        if control_value1[20] == (self.api_DateTimeM):
                                                            self.ws.write(self.rowsize, 24, (self.api_DateTimeM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 24, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 24, 'Empty',
                                                                      self.__style9)

                                                    if self.api_DateTime01M:
                                                        if control_value1[21] == (self.api_DateTime01M):
                                                            self.ws.write(self.rowsize, 25, (self.api_DateTime01M),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 25, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 25, 'Empty',
                                                                      self.__style9)

                                                    if self.api_CollegesM:
                                                        if control_value1[22] == (self.api_CollegesM):
                                                            self.ws.write(self.rowsize, 26, (self.api_CollegesM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 26, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 26, 'Empty',
                                                                      self.__style9)

                                                    if self.api_DegreeM:
                                                        if control_value1[23] == (self.api_DegreeM):
                                                            self.ws.write(self.rowsize, 27, (self.api_DegreeM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 27, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 27, 'Empty',
                                                                      self.__style9)

                                                    if self.api_FieldYouWantM:
                                                        if control_value1[24] == str(self.api_FieldYouWantM):
                                                            self.ws.write(self.rowsize, 28, str(self.api_FieldYouWantM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 28, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 28, 'Empty',
                                                                      self.__style9)

                                                    if self.api_GenderM:
                                                        if control_value1[25] == str(self.api_GenderM):
                                                            self.ws.write(self.rowsize, 29, str(self.api_GenderM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 29, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 29, 'Empty',
                                                                      self.__style9)

                                                    if self.api_ExtraCourseM:
                                                        if control_value1[26] == str(self.api_ExtraCourseM):
                                                            self.ws.write(self.rowsize, 30, str(self.api_ExtraCourseM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 30, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 30, 'Empty',
                                                                      self.__style9)

                                                    if self.api_OtherOptionalFieldM:
                                                        if control_value1[27] == str(
                                                                self.api_OtherOptionalFieldM):
                                                            self.ws.write(self.rowsize, 31,
                                                                          str(self.api_OtherOptionalFieldM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 31, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 31, 'Empty',
                                                                      self.__style9)

                                                    if self.api_AddressPermanentM:
                                                        if control_value1[28] == str(
                                                                self.api_AddressPermanentM):
                                                            self.ws.write(self.rowsize, 32,
                                                                          str(self.api_AddressPermanentM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 32, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 32, 'Empty',
                                                                      self.__style9)

                                                    if self.api_AddressCommunicationM:
                                                        if control_value1[29] == str(
                                                                self.api_AddressCommunicationM):
                                                            self.ws.write(self.rowsize, 33,
                                                                          str(self.api_AddressCommunicationM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 33, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 33, 'Empty',
                                                                      self.__style9)

                                                    if self.api_DateM:
                                                        if control_value1[30] == (self.api_DateM):
                                                            self.ws.write(self.rowsize, 34, (self.api_DateM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 34, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 34, 'Empty',
                                                                      self.__style9)

                                                    if self.api_Date01M:
                                                        if control_value1[31] == (self.api_Date01M):
                                                            self.ws.write(self.rowsize, 35, (self.api_Date01M),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 35, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 35, 'Empty',
                                                                      self.__style9)

                                                    if self.api_CurrentTimeM:
                                                        if control_value1[32] == (self.api_CurrentTimeM):
                                                            self.ws.write(self.rowsize, 36, (self.api_CurrentTimeM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 36, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 36, 'Empty',
                                                                      self.__style9)
                                                    if self.api_SubmitTimeM:
                                                        if control_value1[33] == (self.api_SubmitTimeM):
                                                            self.ws.write(self.rowsize, 37, (self.api_SubmitTimeM),
                                                                          self.__style9)
                                                        else:
                                                            self.ws.write(self.rowsize, 37, 'Fail', self.__style7)
                                                    else:
                                                        self.ws.write(self.rowsize, 37, 'Empty',
                                                                      self.__style9)

                                                    self.rowsize += 1  # Row increment
                                                    ob.wb_Result.save(
                                                        '/home/vinodkumar/PycharmProjects/API_Automation/Output Data'
                                                        '/Pofu/Upload_candidates/Submit_Task.xls')

                        m = m + 1


            cur = conn.cursor()
            print(cur)
            conn.close()

        except Exception as e:
            print(e)
            print("DB- Connection Error - Exception Block")

    def over_status(self):
        self.ws.write(0, 0, 'U_Candidates')
        if self.Expected_success_cases == self.Expected_success_cases:
            self.ws.write(0, 1, 'Pass')
        else:
            self.ws.write(0, 1, 'Fail')

        self.ws.write(0, 3, 'StartTime')
        self.ws.write(0, 4, self.start_time)
        ob.wb_Result.save(
            '/home/vinodkumar/PycharmProjects/API_Automation/Output Data/Pofu/Upload_candidates/Submit_Task.xls')

class SubmitUtility():

    def read_from_excel(self, path):
        wb = xlrd.open_workbook(path)
        sheet1 = wb.sheet_by_index(0)
        excel_data_list = []
        for i in range(1, sheet1.nrows):
            individual_candidate_data = []
            number = i
            rows = sheet1.row_values(number)
            # ----------------------------------------------------------
            # Task Details for submit
            # ----------------------------------------------------------

            if not rows[0]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(str(rows[0]))

            if not rows[1]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(str(rows[1]))

            if not rows[2]:
                individual_candidate_data.append(None)
            else:
                # self.xl_DateTime.append(rows[2])
                DateTime = sheet1.cell_value(rowx=(i), colx=2)
                self.DateTime = datetime.datetime(*xlrd.xldate_as_tuple(DateTime, wb.datemode))
                self.DateTime = self.DateTime.strftime("%d-%b-%Y %H:%M")
                individual_candidate_data.append(self.DateTime)
                print(self.DateTime)

            if not rows[3]:
                individual_candidate_data.append(None)
            else:
                # self.xl_DateTime01.append(str(rows[3]))
                DateTime01 = sheet1.cell_value(rowx= (i), colx=3)
                self.DateTime01 = datetime.datetime(*xlrd.xldate_as_tuple(DateTime01, wb.datemode))
                self.DateTime01 = self.DateTime01.strftime("%d-%b-%Y %H:%M")
                individual_candidate_data.append(self.DateTime01)
            if not rows[4]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(int(rows[4]))

            if not rows[5]:
                individual_candidate_data.append(None)

            else:
                individual_candidate_data.append(int(rows[5]))
            if not rows[6]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[6])
            #
            if not rows[7]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[7])
            #
            if not rows[8]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[8])
            #
            if not rows[9]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[9])
            #
            if not rows[10]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[10])
            #
            if not rows[11]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[11])
            #
            if not rows[12]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[12])
            #
            if not rows[13]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[13])
            #
            if not rows[14]:
                individual_candidate_data.append(None)

            else:
                Date = sheet1.cell_value(rowx=(i), colx=14)
                self.Date = datetime.datetime(*xlrd.xldate_as_tuple(Date, wb.datemode))
                self.Date = self.Date.strftime("%d-%b-%Y %H:%M").replace('00:00', "")
                individual_candidate_data.append(self.Date)
                #print self.xl_Date

            if not rows[15]:
                individual_candidate_data.append(None)

            else:
                Date01 = sheet1.cell_value(rowx=(i), colx=15)
                self.Date01 = datetime.datetime(*xlrd.xldate_as_tuple(Date01, wb.datemode))
                self.Date01 = self.Date01.strftime("%d-%b-%Y %H:%M").replace('00:00', "")
                individual_candidate_data.append(self.Date01)
                #print self.xl_Date01

            if not rows[16]:
                individual_candidate_data.append(None)
            else:
                CurrentTime = sheet1.cell_value(rowx=(i), colx=16)
                self.CurrentTime = datetime.datetime(*xlrd.xldate_as_tuple(CurrentTime, wb.datemode))
                self.CurrentTime = self.CurrentTime.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                individual_candidate_data.append(self.CurrentTime)
                print(self.CurrentTime)

            if not rows[17]:
                individual_candidate_data.append(None)
            else:
                SubmitTime = sheet1.cell_value(rowx=(i), colx=17)
                self.SubmitTime = datetime.datetime(*xlrd.xldate_as_tuple(SubmitTime, wb.datemode))
                self.SubmitTime = self.SubmitTime.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                individual_candidate_data.append(self.SubmitTime)
                print(self.SubmitTime)
            #
            if not rows[18]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[18])
            #
            if not rows[19]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[19])
            #
            if not rows[20]:
                individual_candidate_data.append(None)
            else:
                DateTimeM = sheet1.cell_value(rowx=(i), colx=20)
                self.DateTimeM = datetime.datetime(*xlrd.xldate_as_tuple(DateTimeM, wb.datemode))
                self.DateTimeM = self.DateTimeM.strftime("%d-%b-%Y %H:%M").replace('%H:%M', "")
                individual_candidate_data.append(self.DateTimeM)

            if not rows[21]:
                individual_candidate_data.append(None)
            else:
                DateTime01M = sheet1.cell_value(rowx=(i), colx=21)
                self.DateTime01M = datetime.datetime(*xlrd.xldate_as_tuple(DateTime01M, wb.datemode))
                self.DateTime01M = self.DateTime01M.strftime("%d-%b-%Y %H:%M").replace('%H:%M', "")
                individual_candidate_data.append(self.DateTime01M)

            if not rows[22]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(int(rows[22]))
            #
            if not rows[23]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(int(rows[23]))
            #
            if not rows[24]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[24])
            #
            if not rows[25]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[25])
            #
            if not rows[26]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[26])
            #
            if not rows[27]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[27])
            #
            if not rows[28]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[28])
            if not rows[29]:
                individual_candidate_data.append(None)
            else:
                individual_candidate_data.append(rows[29])

            if not rows[30]:
                individual_candidate_data.append(None)
            else:
                DateM = sheet1.cell_value(rowx=(i), colx=30)
                self.DateM = datetime.datetime(*xlrd.xldate_as_tuple(DateM, wb.datemode))
                self.DateM = self.DateM.strftime("%d-%m-%Y")
                individual_candidate_data.append(self.Date)
                #print self.xl_DateM

            if not rows[31]:
                individual_candidate_data.append(None)
            else:
                Date01M = sheet1.cell_value(rowx=(i), colx=31)
                self.Date01M = datetime.datetime(*xlrd.xldate_as_tuple(Date01M, wb.datemode))
                self.Date01M = self.Date01M.strftime("%d-%m-%Y")
                individual_candidate_data.append(self.Date)
                #print self.xl_Date01M

            if not rows[32]:
                individual_candidate_data.append(None)
            else:
                CurrentTimeM = sheet1.cell_value(rowx=(i), colx=32)
                self.CurrentTimeM = datetime.datetime(*xlrd.xldate_as_tuple(CurrentTimeM, wb.datemode))
                self.CurrentTimeM = self.CurrentTimeM.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                individual_candidate_data.append(self.CurrentTimeM)
                print(self.CurrentTimeM)

            if not rows[33]:
                individual_candidate_data.append(None)
            else:
                SubmitTimeM = sheet1.cell_value(rowx=(i), colx=33)
                self.SubmitTimeM = datetime.datetime(*xlrd.xldate_as_tuple(SubmitTimeM, wb.datemode))
                self.SubmitTimeM = self.SubmitTimeM.strftime("%d-%b-%Y %H:%M").replace('30-Dec-2018', "")
                individual_candidate_data.append(self.SubmitTimeM)
                print(self.SubmitTimeM)
            excel_data_list.append(individual_candidate_data)
        return excel_data_list



            #def submitForm(request,token):
ob = Submit_Task()
ob.excel_data()
total_len = len(ob.xl_FirstName)
ob.submit_task_by_candidate(total_len)
#ob1 = SubmitUtility()
#ob1.read_from_excel('/home/vikas/Desktop/PythonScripts/SubmitTask/InputFile/Submit_Task_04.xls')
#ob.over_status()
