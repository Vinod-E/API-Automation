from hpro_automation import (login, input_paths, api)
import xlrd
import json
import requests


class ResetUser(login.CRPOLogin):

    def __init__(self):

        super(ResetUser, self).__init__()

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

    def read_excel(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['resetUser_Input_sheet'])
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
        update_api = requests.post(api.web_api['Update_user'], headers=self.get_token,
                                   data=json.dumps(request, default=str), verify=False)
        update_api_response = json.loads(update_api.content)
        print(update_api_response)


Object = ResetUser()
Object.read_excel()
Total_count = len(Object.xl_update_user_id)
print("Total count ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.update_user(looping)

        Object.user_dict = {}
        Object.update_userId = {}
        Object.update_message = {}
        Object.success_case_01 = {}
        Object.success_case_02 = {}
