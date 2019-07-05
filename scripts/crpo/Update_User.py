from hpro_automation import (login, input_paths, work_book, output_paths)
import xlrd
import datetime
import json
import requests


class UpdateUser(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):

        self.start_time = str(datetime.datetime.now())

        super(UpdateUser, self).__init__()

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 45)))
        self.Actual_Success_case = []

        self.xl_update_user_id = []
        self.xl_update_TypeOfUser = []
        self.xl_update_username = []
        self.xl_update_uname = []
        self.xl_update_email = []
        self.xl_update_Mobile = []
        self.xl_update_UserRoles = []
        self.xl_update_DepartmentId = []
        self.xl_update_LocationId = []
        self.xl_update_UserBelongsTo = []
        self.xl_update_TimeZoneId = []
        self.xl_update_expected_message = []

        self.user_dict = {}
        self.update_userId = {}
        self.update_message = {}
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.headers = {}

    def excel_headers(self):
        self.main_headers = ['Comparison', 'Status', 'User Id', 'UserName', 'Name', 'Email', 'Location', 'Mobile',
                             'Roles', 'Department', 'TypeofUserId', 'UserBelongs_Id', 'TimeZone', 'Expected_message']
        self.headers_with_style2 = ['Comparison', 'Status']
        self.file_headers_col_row()

    def read_excel(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['updateUser_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[0]:
                self.xl_update_user_id.append(None)
            else:
                self.xl_update_user_id.append(int(rows[0]))

            if not rows[1]:
                self.xl_update_TypeOfUser.append(None)
            else:
                self.xl_update_TypeOfUser.append(int(rows[1]))

            if not rows[2]:
                self.xl_update_username.append(None)
            else:
                self.xl_update_username.append(rows[2])

            if not rows[3]:
                self.xl_update_uname.append(None)
            else:
                self.xl_update_uname.append(rows[3])

            if not rows[4]:
                self.xl_update_email.append(None)
            else:
                self.xl_update_email.append(rows[4])

            if not rows[5]:
                self.xl_update_Mobile.append(0)
            else:
                self.xl_update_Mobile.append(int(rows[5]))

            roles = list(map(int, rows[6].split(',') if isinstance(rows[6], str) else [rows[6]]))
            self.xl_update_UserRoles.append(roles)

            if not rows[7]:
                self.xl_update_DepartmentId.append(None)
            else:
                self.xl_update_DepartmentId.append(int(rows[7]))

            if not rows[8]:
                self.xl_update_LocationId.append(None)
            else:
                self.xl_update_LocationId.append(int(rows[8]))

            if not rows[9]:
                self.xl_update_UserBelongsTo.append(None)
            else:
                self.xl_update_UserBelongsTo.append(int(rows[9]))

            if not rows[10]:
                self.xl_update_TimeZoneId.append(None)
            else:
                self.xl_update_TimeZoneId.append(int(rows[10]))

            if not rows[11]:
                self.xl_update_expected_message.append(None)
            else:
                self.xl_update_expected_message.append(rows[11])

    def update_user(self, loop):

        self.lambda_function('Update_user')
        self.headers['APP-NAME'] = 'crpo'

        request = {
            "UserDetails": {
                "UserName": self.xl_update_username[loop],
                "Name": self.xl_update_uname[loop],
                "Email1": self.xl_update_email[loop],
                "TypeOfUser": self.xl_update_TypeOfUser[loop],
                "DepartmentId": self.xl_update_DepartmentId[loop],
                "Mobile1": self.xl_update_Mobile[loop],
                "LocationId": self.xl_update_LocationId[loop],
                "UserRoles": self.xl_update_UserRoles[loop],
                "UserBelongsTo": self.xl_update_UserBelongsTo[loop],
                "TimeZoneId": self.xl_update_TimeZoneId[loop]
            },
            "UserId": self.xl_update_user_id[loop]
        }
        update_api = requests.post(self.webapi, headers=self.headers,
                                   data=json.dumps(request, default=str), verify=False)
        print(update_api.headers)
        update_api_response = json.loads(update_api.content)
        print(update_api_response)
        status = update_api_response['status']
        error = update_api_response.get('error')
        if status == 'OK':
            self.update_userId = update_api_response.get('UserId')
        if status == 'KO':
            self.update_message = error.get('errorDescription')

    def user_getbyid_details(self, loop):

        self.lambda_function('UserGetByid')
        self.headers['APP-NAME'] = 'crpo'

        get_user_details = requests.get(self.webapi.format(self.xl_update_user_id[loop]),
                                        headers=self.headers)
        print(get_user_details.headers)
        user_details = json.loads(get_user_details.content)
        self.user_dict = user_details['UserDetails']

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)

        self.ws.write(self.rowsize, 2, self.xl_update_user_id[loop] if self.xl_update_user_id[loop] else 'Empty')

        self.ws.write(self.rowsize, 3, self.xl_update_username[loop] if self.xl_update_username[loop] else 'Empty')

        self.ws.write(self.rowsize, 4, self.xl_update_uname[loop] if self.xl_update_uname[loop] else 'Empty')

        self.ws.write(self.rowsize, 5, self.xl_update_email[loop] if self.xl_update_email[loop] else 'Empty')

        self.ws.write(self.rowsize, 6, self.xl_update_LocationId[loop] if self.xl_update_LocationId[loop] else 'Empty')

        self.ws.write(self.rowsize, 7, self.xl_update_Mobile[loop] if self.xl_update_Mobile[loop] else 'Empty')

        self.ws.write(self.rowsize, 8,
                      ','.join(map(str, self.xl_update_UserRoles[loop])) if self.xl_update_UserRoles[loop] else 'Empty')

        self.ws.write(self.rowsize, 9,
                      self.xl_update_DepartmentId[loop] if self.xl_update_DepartmentId[loop] else 'Empty')

        self.ws.write(self.rowsize, 10, self.xl_update_TypeOfUser[loop] if self.xl_update_TypeOfUser[loop] else 'Empty')

        self.ws.write(self.rowsize, 11,
                      self.xl_update_UserBelongsTo[loop] if self.xl_update_UserBelongsTo[loop] else 'Empty')

        self.ws.write(self.rowsize, 12, self.xl_update_TimeZoneId[loop] if self.xl_update_TimeZoneId[loop] else 'Empty')

        self.ws.write(self.rowsize, 13, self.xl_update_expected_message[loop])

        self.rowsize += 1  # Row increment
        # -------------------
        # Writing Output data
        # -------------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)

        if self.update_userId:
            self.ws.write(self.rowsize, 1, 'Pass', self.style26)
            self.success_case_01 = 'Pass'
        elif self.update_message:
            if self.xl_update_expected_message[loop] == self.update_message:
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_02 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------

        if not self.update_message:
            if self.xl_update_user_id[loop] == self.update_userId:
                if self.xl_update_user_id[loop] is None:
                    self.ws.write(self.rowsize, 2, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 2, self.update_userId, self.style14)
            else:
                self.ws.write(self.rowsize, 2, self.update_userId, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_username[loop] == self.user_dict.get('UserName'):
                if self.xl_update_username[loop] is None:
                    self.ws.write(self.rowsize, 3, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 3, self.user_dict.get('UserName'), self.style14)
            else:
                self.ws.write(self.rowsize, 3, self.user_dict.get('UserName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_uname[loop] == self.user_dict.get('Name'):
                if self.xl_update_uname[loop] is None:
                    self.ws.write(self.rowsize, 4, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 4, self.user_dict.get('Name'), self.style14)
            else:
                self.ws.write(self.rowsize, 4, self.user_dict.get('Name'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_email[loop] == self.user_dict.get('Email1'):
                if self.xl_update_email[loop] is None:
                    self.ws.write(self.rowsize, 5, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 5, self.user_dict.get('Email1'), self.style14)
            else:
                self.ws.write(self.rowsize, 5, self.user_dict.get('Email1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_LocationId[loop] == self.user_dict.get('LocationId'):
                if self.xl_update_LocationId[loop] is None:
                    self.ws.write(self.rowsize, 6, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 6, self.user_dict.get('LocationId'), self.style14)
            else:
                self.ws.write(self.rowsize, 6, self.user_dict.get('LocationId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_Mobile[loop] == int(str(self.user_dict.get('Mobile1', 0))):
                if not self.xl_update_Mobile[loop]:
                    self.ws.write(self.rowsize, 7, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 7, int(self.user_dict.get('Mobile1')), self.style14)
            else:
                self.ws.write(self.rowsize, 7, self.user_dict.get('Mobile1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if not self.update_message:
            if self.xl_update_UserRoles[loop].sort() == self.user_dict.get('UserRoles') \
                    .sort() if self.user_dict.get('UserRoles') else None:
                if self.xl_update_UserRoles[loop] is None:
                    self.ws.write(self.rowsize, 8, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 8, ','.join(map(str, self.user_dict.get('UserRoles'))), self.style14)
            else:
                self.ws.write(self.rowsize, 8, ','.join(map(str, self.user_dict.get('UserRoles'))), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_DepartmentId[loop] == self.user_dict.get('DepartmentId'):
                if self.xl_update_DepartmentId[loop] is None:
                    self.ws.write(self.rowsize, 9, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 9, self.user_dict.get('DepartmentId'), self.style14)
            else:
                self.ws.write(self.rowsize, 9, self.user_dict.get('DepartmentId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_TypeOfUser[loop] == self.user_dict.get('TypeOfUser'):
                if self.xl_update_TypeOfUser[loop] is None:
                    self.ws.write(self.rowsize, 10, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 10, self.user_dict.get('TypeOfUser'), self.style14)
            else:
                self.ws.write(self.rowsize, 10, self.user_dict.get('TypeOfUser'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_UserBelongsTo[loop] == self.user_dict.get('UserBelongsToId'):
                if self.xl_update_UserBelongsTo[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 11, self.user_dict.get('UserBelongsToId'), self.style14)
            else:
                self.ws.write(self.rowsize, 11, self.user_dict.get('UserBelongsToId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.update_message:
            if self.xl_update_TimeZoneId[loop] == self.user_dict.get('TimeZoneId'):
                if self.xl_update_TimeZoneId[loop] is None:
                    self.ws.write(self.rowsize, 12, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 12, self.user_dict.get('TimeZoneId'), self.style14)
            else:
                self.ws.write(self.rowsize, 12, self.user_dict.get('TimeZoneId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.update_message:
            if self.xl_update_expected_message[loop]:
                if self.update_message and 'User' in self.xl_update_expected_message[loop]:
                    self.ws.write(self.rowsize, 13, self.update_message, self.style14)
                elif self.update_message and 'Email already' in self.xl_update_expected_message[loop]:
                    self.ws.write(self.rowsize, 13, self.update_message, self.style14)
                else:
                    self.ws.write(self.rowsize, 13, self.update_message, self.style3)
            else:
                self.ws.write(self.rowsize, 13, self.update_message, self.style3)

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['UpdateUser_Output_sheet'])

        # --------------------- overall success cases ------------------------------------
        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)

    def overall_status(self):
        self.ws.write(0, 0, 'Update User', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)
        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['UpdateUser_Output_sheet'])


Object = UpdateUser()
Object.excel_headers()
Object.read_excel()
Total_count = len(Object.xl_update_user_id)
print("Total count ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.update_user(looping)
        Object.user_getbyid_details(looping)
        Object.output_excel(looping)

        Object.user_dict = {}
        Object.update_userId = {}
        Object.update_message = {}
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.headers = {}


Object.overall_status()
