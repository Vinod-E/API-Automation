from hpro_automation import (input_paths, api, output_paths, work_book, login)
import json
import requests
import xlrd
import urllib3
import datetime


class LoginCheck(work_book.WorkBook, login.CommonLogin):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(LoginCheck, self).__init__()
        # self.lambda_call = str(input("Lambda On/Off:: "))

        # --------------------------
        # Initialising the excel data
        # --------------------------
        self.xl_tenant = []
        self.xl_username = []
        self.xl_pwd = []
        self.xl_API_hits = []
        self.xl_password_change = []
        self.xl_expected_message = []
        self.xl_expected_status = []

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 64)))
        self.Actual_Success_case = []

        # --------------------
        # Initializing the dict
        # --------------------
        self.login_check_api_response = {}
        self.error_message = {}

        self.success_case_01 = {}
        self.success_case_02 = {}

        self.excel_headers()

    def excel_headers(self):
        self.main_headers = ['Actual_Status', 'Tenant', 'Username', 'Password', 'Expected_Login_Status',
                             'Actual_Login_Status', 'Expected_Password_change', 'Actual_Password_change',
                             'Expected_Message', 'Actual_Message', 'LastLogin']
        self.headers_with_style2 = ['Tenant', 'Username', 'Password', 'Actual_Status']
        self.file_headers_col_row()

    def excel_data(self):

        # ---------------
        # Read excel data
        # ---------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['Login_check_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0] is not None and rows[0] == '':
                self.xl_tenant.append(None)
            else:
                self.xl_tenant.append(rows[0])

            if rows[1] is not None and rows[1] == '':
                self.xl_username.append(None)
            else:
                self.xl_username.append(rows[1])

            if rows[2] is not None and rows[2] == '':
                self.xl_pwd.append(None)
            else:
                self.xl_pwd.append(rows[2])

            if rows[3] is not None and rows[3] == '':
                self.xl_API_hits.append(None)
            else:
                self.xl_API_hits.append(int(rows[3]))

            if rows[4] is not None and rows[4] == '':
                self.xl_password_change.append(None)
            else:
                self.xl_password_change.append(rows[4])

            if rows[5] is not None and rows[5] == '':
                self.xl_expected_message.append(None)
            else:
                self.xl_expected_message.append(rows[5])

            if rows[6] is not None and rows[6] == '':
                self.xl_expected_status.append(None)
            else:
                self.xl_expected_status.append(rows[6])

    def login_check(self, loop):

        if self.calling_lambda == 'On' or self.calling_lambda == 'on':
            if api.lambda_apis.get('Loginto_CRPO') is not None:
                self.headers = self.lambda_headers
                self.webapi = api.lambda_apis['Loginto_CRPO']
        else:
            self.headers = self.Non_lambda_headers
            self.webapi = api.lambda_apis['Loginto_CRPO']

        urllib3.disable_warnings()
        request = {"LoginName": self.xl_username[loop],
                   "Password": self.xl_pwd[loop],
                   "UserName": self.xl_username[loop],
                   "TenantAlias": self.xl_tenant[loop]
                   }
        # -----------------------------------------------------
        # Hitting login API multiple times based on input value
        # -----------------------------------------------------
        for i in range(self.xl_API_hits[loop]):
            login_check = requests.post(self.webapi, headers=self.headers,
                                        data=json.dumps(request, default=str), verify=False)
            print(login_check.headers)

            self.login_check_api_response = json.loads(login_check.content)
            print(self.login_check_api_response)

            if self.login_check_api_response.get('status') == 'KO':
                error = self.login_check_api_response.get('error')
                self.error_message = error.get('errorDescription')

    def output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        if self.xl_tenant[loop]:
            self.ws.write(self.rowsize, 1, self.xl_tenant[loop])
        else:
            self.ws.write(self.rowsize, 1, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_username[loop]:
            self.ws.write(self.rowsize, 2, self.xl_username[loop])
        else:
            self.ws.write(self.rowsize, 2, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_pwd[loop]:
            self.ws.write(self.rowsize, 3, self.xl_pwd[loop])
        else:
            self.ws.write(self.rowsize, 3, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_expected_status[loop]:
            self.ws.write(self.rowsize, 4, self.xl_expected_status[loop])
        else:
            self.ws.write(self.rowsize, 4, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_password_change[loop]:
            self.ws.write(self.rowsize, 6, self.xl_password_change[loop])
        else:
            self.ws.write(self.rowsize, 6, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_expected_message[loop]:
            self.ws.write(self.rowsize, 8, self.xl_expected_message[loop])
        else:
            self.ws.write(self.rowsize, 8, 'Empty', self.style7)
        # --------------------------------------------------------------------------------------------------------------

        # -------------------
        # Writing Output Data
        # -------------------

        if self.xl_expected_message[loop] == self.error_message:
            self.ws.write(self.rowsize, 0, 'Pass', self.style26)
            self.success_case_01 = 'Pass'
        elif self.login_check_api_response.get('LastLogin'):
            if self.xl_expected_message[loop]:
                self.ws.write(self.rowsize, 0, 'Pass', self.style26)
                self.success_case_02 = 'Pass'
            else:
                self.ws.write(self.rowsize, 0, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 0, 'Fail', self.style3)

        # --------------------------------------------------------------------------------------------------------------

        if self.login_check_api_response:
            if self.login_check_api_response.get('status') == 'OK':
                if self.xl_expected_status[loop] == 'Pass':
                    self.ws.write(self.rowsize, 5, 'Pass', self.style8)
            elif self.login_check_api_response.get('status') == 'KO':
                if self.xl_expected_status[loop] == 'Fail':
                    self.ws.write(self.rowsize, 5, 'Fail', self.style8)
        # --------------------------------------------------------------------------------------------------------------

        if self.login_check_api_response:
            if self.login_check_api_response.get('IsPasswordChangeRequired'):
                if self.xl_password_change[loop] == 'Yes':
                    self.ws.write(self.rowsize, 7, 'Yes', self.style8)
            else:
                if self.xl_password_change[loop] == 'No':
                    self.ws.write(self.rowsize, 7, 'No', self.style8)
        # --------------------------------------------------------------------------------------------------------------

        if self.error_message:
            if self.xl_expected_message[loop] == self.error_message:
                self.ws.write(self.rowsize, 9, self.error_message, self.style8)
            else:
                self.ws.write(self.rowsize, 9, self.error_message, self.style3)

        elif self.login_check_api_response.get('LastLogin'):
            if self.xl_expected_message[loop]:
                if 'Login Suc' in self.xl_expected_message[loop]:
                    self.ws.write(self.rowsize, 9, self.xl_expected_message[loop], self.style8)
        else:
            self.ws.write(self.rowsize, 9,)
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 10, self.login_check_api_response.get('LastLogin'))
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        self.col += 1  # Column increment
        Object.wb_Result.save(output_paths.outputpaths['Login_check_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)

    def overall_status(self):
        self.ws.write(0, 0, 'Login Check', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'StartTime', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Login_check_Output_sheet'])


Object = LoginCheck()
Object.excel_data()
Total_count = len(Object.xl_expected_message)
print("Number of rows::", Total_count)

for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)

    Object.login_check(looping)
    Object.output_excel(looping)

    # -----------------
    # Making dict empty
    # -----------------
    Object.login_check_api_response = {}
    Object.error_message = {}
    Object.success_case_01 = {}
    Object.success_case_02 = {}

Object.overall_status()
