import requests
import json
import unittest
import xlrd

class createUsers(unittest.TestCase):
    def test_createUsers(self):
        headers1 = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "Hirepro#aocm798", "TenantAlias": "wings", "UserName": "admin" }
        response = requests.post("https://ams.hirepro.in/py/common/user/login_user/", headers = headers1, data=json.dumps(data), verify=False)
        self.TokenVal = response.json()
        print self.TokenVal
        wb = xlrd.open_workbook("/home/sanjeev/createUser3.xls")
        sheetname = wb.sheet_names()  # Read for XLS Sheet names
        print(sheetname)
        sh1 = wb.sheet_by_index(0)  # add login details
        i = 1
        rownum = i
        while i < sh1.nrows:
            rows = sh1.row_values(rownum)
            headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
            # data = {"UserDetails": {"UserName": rows[0], "Name": rows[1], "Email1": rows[2],
            #                         "IsPasswordAutoGenerated": False, "Password": rows[3], "TypeOfUser": 1,
                                    # "LocationId": rows[4], "UserRoles": [rows[5]]}}

            # databuddy = {"UserDetails":{"UserName":rows[0],"Name":rows[1],"Email1":rows[2],"IsPasswordAutoGenerated":False,"Password":rows[3],"TypeOfUser":1,"DepartmentId":3383,"LocationId":25187,"UserRoles":[9663],"ReportTo":"","Organization":"","BU":"","SBU":"","PL":""}}
            databuddy = {"UserDetails": {"UserName": rows[0], "Name": rows[1],
                             "Email1": rows[2], "IsPasswordAutoGenerated": False,
                             "Password": rows[3], "TypeOfUser": 1, "DepartmentId": 5495, "LocationId": 25187,
                             "UserRoles": [17898]}}

            r = requests.post("https://ams.hirepro.in/py/common/user/create_user/", headers=headers, data=json.dumps(databuddy, default=str), verify=False)
            statusCode = json.loads(r.content)['statusCode']
            # print r.statusCode
            if statusCode == 200:
                print json.loads(r.content)['UserId']
            else:
                print statusCode
            i += 1
            rownum += 1