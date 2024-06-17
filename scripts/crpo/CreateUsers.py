import requests
import json
import datetime
import xlrd
from hpro_automation import (input_paths, output_paths, login, work_book, db_login)
from hpro_automation.api import *


class CreateUser(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(CreateUser, self).__init__()
        self.common_login('admin')
        self.crpo_app_name = self.app_name.strip()
        print(self.crpo_app_name)
        self.db_connection()

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Typeofuser = []  # [] Initialising data from excel sheet to the variables
        self.xl_Name = []
        self.xl_Login_Name = []
        self.xl_Email = []
        self.xl_Mobile = []
        self.xl_Roles = []
        self.xl_Department = []
        self.xl_Location = []
        self.xl_enter_password = []
        self.xl_UserBelongsTo = []
        self.xl_Execption_Message = []
        self.xl_auto_password = []

        self.db_salt_password = ''

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 23)))
        self.Actual_Success_case = []

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.user_dict = {}
        self.userId = {}
        self.error = {}
        self.message = {}
        self.status = {}
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.status = {}
        self.create_user_request = {}

        self.excel_headers()

    def excel_headers(self):
        self.main_headers = ['Comparison', 'Actual_status', 'User Id', 'TypeofUser', 'Name', 'Login_name',
                             'Email', 'Location', 'Mobile', 'Roles', 'Department', 'TypeofUserId', 'UserBelongs_Id',
                             'BD_Salt_Password', 'Expected_message']
        self.headers_with_style2 = ['Comparison', 'Actual_status', 'User Id', 'TypeofUser']
        self.file_headers_col_row()

    def excel_data(self):

        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['CreateUser_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[0]:
                self.xl_Typeofuser.append(None)
            else:
                self.xl_Typeofuser.append(int(rows[0]))

            if not rows[1]:
                self.xl_Name.append(None)
            else:
                self.xl_Name.append(str(rows[1]))

            if not rows[2]:
                self.xl_Login_Name.append(None)
            else:
                self.xl_Login_Name.append(str(rows[2]))

            if not rows[3]:
                self.xl_Email.append(None)
            else:
                self.xl_Email.append(str(rows[3]))

            if not rows[4]:
                self.xl_Mobile.append(None)
            else:
                self.xl_Mobile.append(int(rows[4]))

            if not rows[6]:
                self.xl_Department.append(None)
            else:
                self.xl_Department.append(int(rows[6]))

            if not rows[7]:
                self.xl_Location.append(None)
            else:
                self.xl_Location.append(int(rows[7]))

            if not rows[8]:
                self.xl_enter_password.append(None)
            else:
                self.xl_enter_password.append(str(rows[8]))

            if not rows[9]:
                self.xl_UserBelongsTo.append(None)
            else:
                self.xl_UserBelongsTo.append(int(rows[9]))

            if not rows[10]:
                self.xl_Execption_Message.append(None)
            else:
                self.xl_Execption_Message.append(str(rows[10]))

            if not rows[11]:
                self.xl_auto_password.append(None)
            else:
                self.xl_auto_password.append(rows[11])

            roles = list(map(int, rows[5].split(',') if isinstance(rows[5], str) else [rows[5]]))
            self.xl_Roles.append(roles)

    def update_pwd_policy(self):

        self.lambda_function('create_update_pwd_policy')
        self.headers['APP-NAME'] = self.crpo_app_name

        request = {"NumCapital": 1,
                   "NumSmall": 1,
                   "NumSpecial": 1,
                   "NumNumeric": 1,
                   "NumCharacter": 4,
                   "PwdExpiryLimitInDays": 60,
                   "IsPwdChangeInFirstLogin": False,
                   "Id": 519
                   }
        update_policy = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                      verify=False)
        print(update_policy.headers)

    def remove_tenant_cache(self):
        self.lambda_function('tenant_cache')
        request = {}
        tenant_cache = requests.post(self.webapi, headers=self.headers, data=json.dumps(request, default=str),
                                     verify=False)
        remove_tenant_cache = json.loads(tenant_cache.content)
        print(remove_tenant_cache)

    def create_user(self, loop):

        self.lambda_function('Create_user')
        self.headers['APP-NAME'] = self.crpo_app_name

        if self.xl_auto_password[loop] == 1:

            self.create_user_request = {
                "UserDetails": {"Name": self.xl_Name[loop], "UserName": self.xl_Login_Name[loop],
                                "Email1": self.xl_Email[loop], "Password": self.xl_enter_password[loop],
                                "IsPasswordAutoGenerated": False, "TypeOfUser": self.xl_Typeofuser[loop],
                                "Mobile1": self.xl_Mobile[loop], "LocationId": self.xl_Location[loop],
                                "UserRoles": self.xl_Roles[loop], "DepartmentId": self.xl_Department[loop],
                                "UserBelongsTo": self.xl_UserBelongsTo[loop]}
            }
        else:
            self.create_user_request = {
                "UserDetails": {"Name": self.xl_Name[loop], "UserName": self.xl_Login_Name[loop],
                                "Email1": self.xl_Email[loop], "IsPasswordAutoGenerated": True,
                                "TypeOfUser": self.xl_Typeofuser[loop], "Mobile1": self.xl_Mobile[loop],
                                "LocationId": self.xl_Location[loop], "UserRoles": self.xl_Roles[loop],
                                "DepartmentId": self.xl_Department[loop], "UserBelongsTo": self.xl_UserBelongsTo[loop]}
            }
        create_user = requests.post(self.webapi, headers=self.headers,
                                    data=json.dumps(self.create_user_request, default=str), verify=False)
        print(create_user.headers)
        create_user_response = json.loads(create_user.content)
        print(create_user_response)
        self.status = create_user_response['status']
        self.userId = create_user_response.get('UserId')
        self.error = create_user_response.get('error', {})
        self.message = self.error.get('errorDescription')
        if self.status == 'OK':
            print("Create User successfully")
            print("Status is", self.status)
        else:
            print("user has not been created")
            print("Status is", self.status)

        # -------------- query for password ------------------------
        if self.userId:
            user_query = "select salt_password from users where id = {}".format(self.userId)
            self.cursor.execute(user_query)
            print(user_query)
            self.cursor.execute(user_query)
            records = self.cursor.fetchall()
            self.connection.commit()
            print("Total number of rows ::", self.cursor.rowcount)
            for row in records:
                if row[0] is not None:
                    self.db_salt_password = str(row[0])
                else:
                    self.db_salt_password = str(row[0])

    def user_getbyid_details(self):

        self.lambda_function('UserGetByid')
        self.headers['APP-NAME'] = self.crpo_app_name

        get_user_details = requests.get(self.webapi.format(self.userId), headers=self.headers)
        print(get_user_details.headers)
        user_details = json.loads(get_user_details.content)
        self.user_dict = user_details.get('UserDetails')

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        if self.xl_Name[loop]:
            self.ws.write(self.rowsize, 4, self.xl_Name[loop])
        else:
            self.ws.write(self.rowsize, 4, 'Empty')
        if self.xl_Login_Name[loop]:
            self.ws.write(self.rowsize, 5, self.xl_Login_Name[loop])
        else:
            self.ws.write(self.rowsize, 5, 'Empty')
        if self.xl_Email[loop]:
            self.ws.write(self.rowsize, 6, self.xl_Email[loop])
        else:
            self.ws.write(self.rowsize, 6, 'Empty')
        if self.xl_Location[loop]:
            self.ws.write(self.rowsize, 7, self.xl_Location[loop])
        else:
            self.ws.write(self.rowsize, 7, 'Empty')
        if self.xl_Mobile[loop]:
            self.ws.write(self.rowsize, 8, self.xl_Mobile[loop])
        else:
            self.ws.write(self.rowsize, 8, 'Empty')
        if self.xl_Roles[loop]:
            self.ws.write(self.rowsize, 9, ','.join(map(str, self.xl_Roles[loop])))
        else:
            self.ws.write(self.rowsize, 9, 'Empty')
        if self.xl_Department[loop]:
            self.ws.write(self.rowsize, 10, self.xl_Department[loop])
        else:
            self.ws.write(self.rowsize, 10, 'Empty')
        if self.xl_Typeofuser[loop]:
            self.ws.write(self.rowsize, 11, self.xl_Typeofuser[loop])
        else:
            self.ws.write(self.rowsize, 11, 'Empty')
        if self.xl_UserBelongsTo[loop]:
            self.ws.write(self.rowsize, 12, self.xl_UserBelongsTo[loop])
        else:
            self.ws.write(self.rowsize, 12, 'Empty')
        if self.xl_Execption_Message[loop]:
            self.ws.write(self.rowsize, 14, self.xl_Execption_Message[loop])
        if self.xl_enter_password[loop]:
            self.ws.write(self.rowsize, 13, self.xl_enter_password[loop])
        else:
            self.ws.write(self.rowsize, 13, 'Empty')

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        if self.userId:
            if self.xl_Execption_Message[loop] == '':
                self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                self.success_case_01 = 'Pass'
            elif self.xl_Execption_Message[loop] == self.message:
                self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                self.success_case_01 = 'Pass'
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        elif self.xl_Execption_Message[loop] == self.message:
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
            self.success_case_02 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)

        self.ws.write(self.rowsize, 2, self.userId)
        self.ws.write(self.rowsize, 3, self.user_dict.get('TypeOfUserText'))

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------

        if self.userId:
            if self.xl_Name[loop] == self.user_dict.get('Name'):
                if self.xl_Name[loop] is None:
                    self.ws.write(self.rowsize, 4, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 4, self.user_dict.get('Name'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 4, self.user_dict.get('Name'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 4, self.user_dict.get('Name'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Login_Name[loop] == self.user_dict.get('UserName'):
                if self.xl_Login_Name[loop] is None:
                    self.ws.write(self.rowsize, 5, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 5, self.user_dict.get('UserName'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 5, self.user_dict.get('UserName'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 5, self.user_dict.get('UserName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Email[loop] == self.user_dict.get('Email1'):
                if self.xl_Email[loop] is None:
                    self.ws.write(self.rowsize, 6, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 6, self.user_dict.get('Email1'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 6, self.user_dict.get('Email1'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 6, self.user_dict.get('Email1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Location[loop] == self.user_dict.get('LocationId'):
                if self.xl_Location[loop] is None:
                    self.ws.write(self.rowsize, 7, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 7, self.user_dict.get('LocationId'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 7, self.user_dict.get('LocationId'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 7, self.user_dict.get('LocationId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Mobile[loop] == int(self.user_dict.get('Mobile1', 0) if self.user_dict.get('Mobile1') else 0):
                if self.xl_Mobile[loop] is None:
                    self.ws.write(self.rowsize, 8, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 8, self.user_dict.get('Mobile1'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 8, self.user_dict.get('Mobile1'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 8, self.user_dict.get('Mobile1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Roles[loop].sort() == self.user_dict.get('UserRoles') \
                    .sort() if self.user_dict.get('UserRoles') else None:
                if self.xl_Roles[loop] is None:
                    self.ws.write(self.rowsize, 9,
                                  self.message if self.message else 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 9, ','.join(map(str, self.user_dict.get('UserRoles'))), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 9, self.user_dict.get('UserRoles'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 9, self.user_dict.get('UserRoles'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Department[loop] == self.user_dict.get('DepartmentId'):
                if self.xl_Department[loop] is None:
                    self.ws.write(self.rowsize, 10, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 10, self.user_dict.get('DepartmentId'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 10, self.user_dict.get('DepartmentId'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 10, self.user_dict.get('DepartmentId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.userId:
            if self.xl_Typeofuser[loop] == self.user_dict.get('TypeOfUser'):
                if self.xl_Typeofuser[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 11, self.user_dict.get('TypeOfUser'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 11, self.user_dict.get('TypeOfUser'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 11, self.user_dict.get('TypeOfUser'), self.style3)
                else:
                    self.ws.write(self.rowsize, 11, self.user_dict.get('TypeOfUser'), self.style6)
        # --------------------------------------------------------------------------------------------------------------
        if self.userId:
            if self.xl_UserBelongsTo[loop] == self.user_dict.get('UserBelongsToId'):
                if self.xl_UserBelongsTo[loop] is None:
                    self.ws.write(self.rowsize, 12, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 12, self.user_dict.get('UserBelongsToId'), self.style14)
            elif self.message == self.message:
                if self.message and 'User' in self.message:
                    self.ws.write(self.rowsize, 12, self.user_dict.get('UserBelongsToId'), self.style3)
                elif self.message and 'Email already' in self.message:
                    self.ws.write(self.rowsize, 12, self.user_dict.get('UserBelongsToId'), self.style3)
                else:
                    self.ws.write(self.rowsize, 12, self.user_dict.get('UserBelongsToId'), self.style6)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_Execption_Message[loop] == self.message:
            if self.xl_Execption_Message[loop] is None:
                self.ws.write(self.rowsize, 14, '')
            else:
                self.ws.write(self.rowsize, 14, self.message, self.style14)
        elif self.message == self.message:
            if self.message and 'User' in self.message:
                self.ws.write(self.rowsize, 14, self.user_dict.get('errorDescription'), self.style3)
            elif self.message and 'Email already' in self.message:
                self.ws.write(self.rowsize, 14, self.user_dict.get('errorDescription'), self.style3)
            elif self.status == 'OK':
                self.ws.write(self.rowsize, 14, 'Create User successfully', self.style3)
            else:
                self.ws.write(self.rowsize, 14, self.user_dict.get('errorDescription'),
                              self.style6)
        else:
            self.ws.write(self.rowsize, 14, self.user_dict.get('errorDescription'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.db_salt_password:
            self.ws.write(self.rowsize, 13, self.db_salt_password, self.style6)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Obj.wb_Result.save(output_paths.outputpaths['CreateUser_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)

    def overall_status(self):
        self.ws.write(0, 0, 'Create User', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)
        self.ws.write(0, 2, 'Login Server', self.style23)
        self.ws.write(0, 3, login_server, self.style24)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'APP Name', self.style23)
        self.ws.write(0, 7, self.crpo_app_name, self.style24)
        self.ws.write(0, 8, 'No.of Test cases', self.style23)
        self.ws.write(0, 9, Total_count, self.style24)
        self.ws.write(0, 10, 'Start Time', self.style23)
        self.ws.write(0, 11, self.start_time, self.style26)
        Obj.wb_Result.save(output_paths.outputpaths['CreateUser_Output_sheet'])


Obj = CreateUser()
Obj.excel_data()
Obj.update_pwd_policy()
Obj.remove_tenant_cache()
Total_count = len(Obj.xl_Name)
print("Number Of Rows ::", Total_count)
if Obj.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Obj.create_user(looping)
        if Obj.status == 'OK':
            Obj.user_getbyid_details()
        Obj.output_excel(looping)

        # ------------------
        # Making Dict empty
        # ------------------
        Obj.userId = {}
        Obj.status = {}
        Obj.user_dict = {}
        Obj.error = {}
        Obj.message = {}
        Obj.status = {}
        Obj.success_case_01 = {}
        Obj.success_case_02 = {}
        Obj.create_user_request = {}
        Obj.db_salt_password = {}
Obj.overall_status()
Obj.connection.close()
