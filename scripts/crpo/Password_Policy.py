from hpro_automation import (login, input_paths, output_paths, db_login, work_book)
import xlrd
import requests
import json
import datetime


class PasswordPolicy(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(PasswordPolicy, self).__init__()
        self.common_login('crpo')

        # ----------------------------------------------
        # Password configurations data set initialisation
        # ----------------------------------------------
        self.xl_ID = []
        self.xl_ExpierDays = []
        self.xl_firstlogin = []
        self.xl_TotalCharacters = []
        self.xl_Numeric = []
        self.xl_capital = []
        self.xl_small = []
        self.xl_special = []

        self.xl_username = []
        self.xl_old_pwd = []
        self.xl_new_pwd = []
        self.xl_confirm_pwd = []

        self.xl_ExceptionMessage = []
        self.xl_p_p_status = []
        self.xl_change_status = []
        self.xl_login_status = []

        self.db_totalcharacters = []
        self.db_capital = []
        self.db_small = []
        self.db_special = []
        self.db_numeric = []

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 29)))
        self.Actual_Success_case = []

        # -----------
        # Dictionary
        # -----------
        self.policy_status = {}
        self.policyid = {}
        self.change_pwd_status = {}
        self.change_pwd_error = {}
        self.login_check_api_response = {}
        self.P_P_dict = {}

        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}

    def excel_headers(self):
        self.main_headers = ['Comparision', 'Overall Status',
                             'TotalCharacters', 'Capital', 'Small', 'Special', 'Numeric', 'ID', 'UserName', 'Password',
                             'LastLogin', 'Pwd_configuration_status', 'Change_Pwd_Status', 'Login_Status',
                             'Exception Error']
        self.headers_with_style2 = ['Comparision', 'Overall Status']
        self.headers_with_style19 = ['TotalCharacters', 'Capital', 'Small', 'Special', 'Numeric', 'ID']
        self.file_headers_col_row()

    def excel_data(self):

        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['PasswordPolicy_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0] is None and rows[0] != '':
                self.xl_ID.append(None)
            else:
                self.xl_ID.append(int(rows[0]))

            if rows[1] is None and rows[1] != '':
                self.xl_ExpierDays.append(None)
            else:
                self.xl_ExpierDays.append(int(rows[1]))

            if rows[2] is None and rows[2] != '':
                self.xl_TotalCharacters.append(None)
            else:
                self.xl_TotalCharacters.append(int(rows[2]))

            if rows[3] is None and rows[3] != '':
                self.xl_Numeric.append(None)
            else:
                self.xl_Numeric.append(int(rows[3]))

            if rows[4] is None and rows[4] != '':
                self.xl_special.append(None)
            else:
                self.xl_special.append(int(rows[4]))

            if rows[5] is None and rows[5] != '':
                self.xl_small.append(None)
            else:
                self.xl_small.append(int(rows[5]))

            if rows[6] is None and rows[6] != '':
                self.xl_capital.append(None)
            else:
                self.xl_capital.append(int(rows[6]))

            if rows[7] is None and rows[7] != '':
                self.xl_username.append(None)
            else:
                self.xl_username.append(rows[7])

            if rows[8] is None and rows[8] != '':
                self.xl_old_pwd.append(None)
            else:
                self.xl_old_pwd.append(rows[8])

            if rows[9] is None and rows[9] != '':
                self.xl_new_pwd.append(None)
            else:
                self.xl_new_pwd.append(rows[9])

            if rows[10] is None and rows[10] != '':
                self.xl_confirm_pwd.append(None)
            else:
                self.xl_confirm_pwd.append(rows[10])

            if rows[11] is None and rows[11] != '':
                self.xl_ExceptionMessage.append(None)
            else:
                self.xl_ExceptionMessage.append(rows[11])

            if rows[12] is None and rows[12] != '':
                self.xl_p_p_status.append(None)
            else:
                self.xl_p_p_status.append(rows[12])

            if rows[13] is None and rows[13] != '':
                self.xl_change_status.append(None)
            else:
                self.xl_change_status.append(rows[13])

            if rows[14] is None and rows[14] != '':
                self.xl_login_status.append(None)
            else:
                self.xl_login_status.append(rows[14])

    def update_pwd_policy(self, loop):

        self.lambda_function('create_update_pwd_policy')
        self.headers['APP-NAME'] = 'crpo'

        request = {"NumCapital": self.xl_capital[loop],
                   "NumSmall": self.xl_small[loop],
                   "NumSpecial": self.xl_special[loop],
                   "NumNumeric": self.xl_Numeric[loop],
                   "NumCharacter": self.xl_TotalCharacters[loop],
                   "PwdExpiryLimitInDays": self.xl_ExpierDays[loop],
                   "IsPwdChangeInFirstLogin": False,
                   "Id": self.xl_ID[loop]
                   }
        update_policy = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                      verify=False)
        print(update_policy.headers)
        update_policy_api_response = json.loads(update_policy.content)
        print(update_policy_api_response)

        self.policy_status = update_policy_api_response['status']
        self.policyid = update_policy_api_response.get('PwdPolicyId')

    def db_pwd_policy(self, loop):

        self.db_connection('amsin1')
        query = 'select num_character,num_capital,num_small,num_special,num_numeric ' \
                'from pwd_policy_configurations where id={};'.format(self.xl_ID[loop])
        self.cursor.execute(query)
        records = self.cursor.fetchall()
        for row in records:
            self.db_totalcharacters = row[0]
            self.db_capital = row[1]
            self.db_small = row[2]
            self.db_special = row[3]
            self.db_numeric = row[4]

            print(self.db_totalcharacters)
        self.cursor.close()

    def change_password(self, loop):

        self.lambda_function('change_password')
        self.headers['APP-NAME'] = 'crpo'

        request = {
            "UserName": self.xl_username[loop],
            "OldPwd": self.xl_old_pwd[loop],
            "NewPwd": self.xl_new_pwd[loop],
            "ConfirmNewPwd": self.xl_confirm_pwd[loop],
            "TenantAlias": "automation"}
        change_password = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                        verify=False)
        print(change_password.headers)
        change_password_api_response = json.loads(change_password.content)
        print(change_password_api_response)
        self.change_pwd_status = change_password_api_response['status']
        self.change_pwd_error = change_password_api_response.get('error')

    def login_check(self, loop):

        self.lambda_function('Loginto_CRPO')
        self.headers['APP-NAME'] = 'crpo'

        request = {"LoginName": self.xl_username[loop],
                   "Password": self.xl_new_pwd[loop],
                   "UserName": self.xl_username[loop],
                   "TenantAlias": 'automation'
                   }
        login_check = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                    verify=False)
        print(login_check.headers)
        self.login_check_api_response = json.loads(login_check.content)
        self.P_P_dict = self.login_check_api_response.get('PasswordPolicy')
        print(self.P_P_dict.get('NumCharacter'))

    def output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 2, self.xl_TotalCharacters[loop])
        self.ws.write(self.rowsize, 3, self.xl_capital[loop])
        self.ws.write(self.rowsize, 4, self.xl_small[loop])
        self.ws.write(self.rowsize, 5, self.xl_special[loop])
        self.ws.write(self.rowsize, 6, self.xl_Numeric[loop])
        self.ws.write(self.rowsize, 7, self.xl_ID[loop])
        self.ws.write(self.rowsize, 8, self.xl_username[loop])
        self.ws.write(self.rowsize, 9, self.xl_confirm_pwd[loop])
        self.ws.write(self.rowsize, 11, self.xl_p_p_status[loop])
        self.ws.write(self.rowsize, 12, self.xl_change_status[loop])
        self.ws.write(self.rowsize, 13, self.xl_login_status[loop])
        self.ws.write(self.rowsize, 14, self.xl_ExceptionMessage[loop])

        # -------------------
        # Writing Output Data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_ID[loop] == self.policyid:
            if self.change_pwd_status == 'OK':
                if self.login_check_api_response.get('status') == 'OK':
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_01 = 'Pass'
                elif self.xl_ExceptionMessage[loop] is not None:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_02 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            elif self.xl_ExceptionMessage[loop] == self.change_pwd_error['errorDescription']:
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_03 = 'Pass'
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.policy_status == 'OK':
            self.ws.write(self.rowsize, 11, 'Allowed', self.style14)
        else:
            self.ws.write(self.rowsize, 11, 'Not Allowed', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.change_pwd_status == 'OK':
            self.ws.write(self.rowsize, 12, 'Allowed', self.style14)
        elif self.xl_ExceptionMessage[loop] and 'policy, please verify' in self.change_pwd_error['errorDescription']:
            if self.xl_login_status[loop] == 'Not Allowed':
                self.ws.write(self.rowsize, 12, 'Not Allowed', self.style14)
            else:
                self.ws.write(self.rowsize, 12, 'Allowed', self.style14)
        elif self.xl_ExceptionMessage[loop] and 'New password and' in self.change_pwd_error['errorDescription']:
            if self.xl_login_status[loop] == 'Not Allowed':
                self.ws.write(self.rowsize, 12, 'Not Allowed', self.style14)
            else:
                self.ws.write(self.rowsize, 12, 'Allowed', self.style14)
        elif self.xl_ExceptionMessage[loop] and 'New password and' in self.change_pwd_error['errorDescription']:
            if self.xl_login_status[loop] == 'Not Allowed':
                self.ws.write(self.rowsize, 12, 'Not Allowed', self.style14)
            else:
                self.ws.write(self.rowsize, 12, 'Allowed', self.style14)
        else:
            self.ws.write(self.rowsize, 12, 'Not Allowed', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.login_check_api_response.get('status') == 'OK':
            self.ws.write(self.rowsize, 13, 'Allowed', self.style14)
        elif self.xl_login_status[loop] == 'Not Allowed':
            self.ws.write(self.rowsize, 13, 'Not Allowed', self.style14)
        else:
            self.ws.write(self.rowsize, 13, 'Not Allowed', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.P_P_dict.get('NumCharacter'):
            if self.P_P_dict.get('NumCharacter') in [0, 8]:
                if self.xl_ExceptionMessage[loop] and 'takes from backend code' in self.xl_ExceptionMessage[loop]:
                    self.ws.write(self.rowsize, 2, self.P_P_dict.get('NumCharacter'), self.style7)
            else:
                self.ws.write(self.rowsize, 2, self.P_P_dict.get('NumCharacter'), self.style14)

        elif self.db_totalcharacters in [0, 4, 8, 9, 64]:
            if self.db_totalcharacters in [0, 8]:
                if self.xl_ExceptionMessage[loop] and 'takes from backend code' in self.xl_ExceptionMessage[loop]:
                    self.ws.write(self.rowsize, 2, 8, self.style7)
            else:
                self.ws.write(self.rowsize, 2, self.db_totalcharacters, self.style14)
        else:
            self.ws.write(self.rowsize, 2, self.db_totalcharacters, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.P_P_dict.get('NumCapital'):
            if self.P_P_dict.get('NumCapital') == self.xl_capital[loop]:
                self.ws.write(self.rowsize, 3, self.P_P_dict.get('NumCapital'), self.style14)
            else:
                self.ws.write(self.rowsize, 3, self.P_P_dict.get('NumCapital'), self.style3)
        elif self.db_capital == self.xl_capital[loop]:
            self.ws.write(self.rowsize, 3, self.db_capital, self.style14)
        else:
            self.ws.write(self.rowsize, 3, self.db_capital, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.P_P_dict.get('NumSmall'):
            if self.P_P_dict.get('NumSmall') == self.xl_small[loop]:
                self.ws.write(self.rowsize, 4, self.P_P_dict.get('NumSmall'), self.style14)
            else:
                self.ws.write(self.rowsize, 4, self.P_P_dict.get('NumSmall'), self.style3)
        elif self.db_small == self.xl_small[loop]:
            self.ws.write(self.rowsize, 4, self.db_small, self.style14)
        else:
            self.ws.write(self.rowsize, 4, self.db_small, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.P_P_dict.get('NumSpecial'):
            if self.P_P_dict.get('NumSpecial') == self.xl_special[loop]:
                self.ws.write(self.rowsize, 5, self.P_P_dict.get('NumSpecial'), self.style14)
            else:
                self.ws.write(self.rowsize, 5, self.P_P_dict.get('NumSpecial'), self.style3)
        elif self.db_special == self.xl_special[loop]:
            self.ws.write(self.rowsize, 5, self.db_special, self.style14)
        else:
            self.ws.write(self.rowsize, 5, self.db_special, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.P_P_dict.get('NumNumeric'):
            if self.P_P_dict.get('NumNumeric') == self.xl_Numeric[loop]:
                self.ws.write(self.rowsize, 6, self.P_P_dict.get('NumNumeric'), self.style14)
            else:
                self.ws.write(self.rowsize, 6, self.P_P_dict.get('NumNumeric'), self.style3)
        elif self.db_numeric == self.xl_Numeric[loop]:
            self.ws.write(self.rowsize, 6, self.db_numeric, self.style14)
        else:
            self.ws.write(self.rowsize, 6, self.db_numeric, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_ID[loop] == self.policyid:
            self.ws.write(self.rowsize, 7, self.policyid, self.style14)
        else:
            self.ws.write(self.rowsize, 7, self.policyid, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_username[loop] == self.login_check_api_response.get('LoginName'):
            self.ws.write(self.rowsize, 8, self.login_check_api_response.get('LoginName'), self.style14)
        else:
            self.ws.write(self.rowsize, 8, self.login_check_api_response.get('LoginName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.login_check_api_response.get('LastLogin'):
            self.ws.write(self.rowsize, 10, self.login_check_api_response.get('LastLogin'))
        else:
            self.ws.write(self.rowsize, 10, self.login_check_api_response.get('LastLogin'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.policy_status != 'OK':
            self.ws.write(self.rowsize, 14, self.change_pwd_error['errorDescription'], self.style3)
        elif self.change_pwd_status != 'OK':
            if self.xl_ExceptionMessage[loop] and 'policy, please verify' in self.change_pwd_error['errorDescription']:
                self.ws.write(self.rowsize, 14, self.change_pwd_error['errorDescription'], self.style14)
            elif self.xl_ExceptionMessage[loop] and 'New password and' in self.change_pwd_error['errorDescription']:
                self.ws.write(self.rowsize, 14, self.change_pwd_error['errorDescription'], self.style14)
            else:
                self.ws.write(self.rowsize, 14, self.change_pwd_error['errorDescription'], self.style3)
        elif self.login_check_api_response['status'] != 'OK':
            self.ws.write(self.rowsize, 14, self.change_pwd_error['errorDescription'], self.style3)
        elif self.xl_ExceptionMessage[loop] and 'takes from backend code' in self.xl_ExceptionMessage[loop]:
            self.ws.write(self.rowsize, 14, self.xl_ExceptionMessage[loop], self.style14)
        elif self.xl_ExceptionMessage[loop] and 'UI validation stops' in self.xl_ExceptionMessage[loop]:
            self.ws.write(self.rowsize, 14, self.xl_ExceptionMessage[loop], self.style14)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Password_policy'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def over_status(self):
        self.ws.write(0, 0, 'Password Policy', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'StartTime', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, Total_count, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Password_policy'])


Object = PasswordPolicy()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_ID)
print("Number of rows ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.update_pwd_policy(looping)
        Object.db_pwd_policy(looping)
        Object.change_password(looping)
        if Object.change_pwd_status == 'OK':
            Object.login_check(looping)
        Object.output_excel(looping)

        # ----------------------------------------
        # Making all dicts are empty for each loop
        # ----------------------------------------
        Object.policyid = {}
        Object.policy_status = {}
        Object.change_pwd_status = {}
        Object.change_pwd_error = {}
        Object.login_check_api_response = {}
        Object.P_P_dict = {}
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.success_case_03 = {}
Object.over_status()
