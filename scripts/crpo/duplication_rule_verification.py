import requests
import json
from hpro_automation.read_excel import *
import datetime
import xlrd
import time
from hpro_automation import (login, input_paths, output_paths, work_book)


class VerifyDuplicationRule(login.CommonLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(VerifyDuplicationRule, self).__init__()
        self.common_login('crpo')

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 92)))
        self.Actual_Success_case = []
        self.success_case_01 = {}

        # print self.headers
        self.excel_headers()
        file_path = input_paths.inputpaths['Duplication_rule_Input_sheet']
        duplicate_sheet_index = 0
        excel_read_obj.excel_read(file_path, duplicate_sheet_index)
        print(excel_read_obj.complete_excel_data)
        data = excel_read_obj.complete_excel_data
        self.tot = len(data)
        print(self.tot)
        for iteration in range(0, self.tot):
            self.current_data = data[iteration]
            self.updateduplicaterule()
            self.checkDuplicate()

    def excel_headers(self):
        self.main_headers = ['Actual_Status', 'Name', 'Fname', 'Mname', 'Lname', 'Email Address', 'Mobile', 'Phone',
                             'Marital Status', 'Gender', 'DOB', 'PANCARD', 'PASSPORT', 'Aadhar', 'USN', 'College',
                             'Degree', 'Location', 'Total Experience(in Months)', 'LinkedIn', 'Facebook', 'Twitter',
                             'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7', 'Text8', 'Text9', 'Text10',
                             'Text11', 'Text12', 'Text13', 'Text14', 'Text15', 'Integer1', 'Integer2', 'Integer3',
                             'Integer4', 'Integer5', 'Integer6', 'Integer7', 'Integer8', 'Integer9', 'Integer10',
                             'Integer11', 'Integer12', 'Integer13', 'Integer14',  'Integer15', 'Duplicate Rule',
                             'Expected Status', 'Actual Status', 'Expected Message', 'Actual Message']
        self.headers_with_style2 = ['Actual_Status', 'Duplicate Rule', 'Expected Status', 'Actual Status',
                                    'Expected Message', 'Actual Message']
        self.file_headers_col_row()

    def updateduplicaterule(self):

        self.lambda_function('save_app_preferences')
        self.headers['APP-NAME'] = 'crpo'

        self.update_json_data = self.current_data.get('DuplicationRuleJson')
        # print self.update_json_data
        self.data1 = {"AppPreference": {"Id": 3595, "Content": self.update_json_data,
                                        "Type": "duplication_conf.default"}, "IsTenantGlobal": "true"}
        r = requests.post(self.webapi, headers=self.headers, data=json.dumps(self.data1, default=str), verify=False)
        print(r.headers)

    def checkDuplicate(self):
        convert_date_of_birth = self.current_data.get('DateOfBirth')
        self.date_of_birth = datetime.datetime(
            *xlrd.xldate_as_tuple(convert_date_of_birth, excel_read_obj.excel_file.datemode))
        self.date_of_birth = self.date_of_birth.strftime("%Y-%m-%d")

        self.data = {"FirstName": self.current_data.get('FirstName'), "MiddleName": self.current_data.get('MiddleName'),
                     "LastName": self.current_data.get('LastName'),
                     "Email1": self.current_data.get('EmailAddress'),
                     "Mobile1": int(self.current_data.get('MobileNumber')) if self.current_data.get(
                         'MobileNumber') else None,
                     "PhoneOffice": int(self.current_data.get('PhoneNumber')) if self.current_data.get(
                         'PhoneNumber') else None,
                     "MaritalStatus": int(self.current_data.get('MaritalStatus')) if self.current_data.get(
                         'MaritalStatus') else None, "Gender": int(self.current_data.get('Gender')),
                     "DateOfBirth": self.date_of_birth,
                     "PanNo": self.current_data.get('Pancard'), "PassportNo": self.current_data.get('Passport'),
                     "AadhaarNo": int(self.current_data.get('Aadhar')) if self.current_data.get('Aadhar') else None,
                     "CollegeId": int(self.current_data.get('College')) if self.current_data.get('College') else None,
                     "DegreeId": int(self.current_data.get('Degree')) if self.current_data.get('Degree') else None,
                     "USN": self.current_data.get('USN'),
                     "CurrentLocationId": int(self.current_data.get('Location')) if self.current_data.get(
                         'Location') else None,
                     "TotalExperience": int(self.current_data.get('TotalExperienceInMonths')) if self.current_data.get(
                         'TotalExperienceInMonths') else None,
                     "FacebookLink": self.current_data.get('Facebook'),
                     "TwitterLink": self.current_data.get('Twitter'),
                     "LinkedInLink": self.current_data.get('LinkedIn'),
                     "Integer1": int(self.current_data.get('Integer1')) if self.current_data.get('Integer1') else None,
                     "Integer2": int(self.current_data.get('Integer2')) if self.current_data.get('Integer2') else None,
                     "Integer3": int(self.current_data.get('Integer3')) if self.current_data.get('Integer3') else None,
                     "Integer4": int(self.current_data.get('Integer4')) if self.current_data.get('Integer4') else None,
                     "Integer5": int(self.current_data.get('Integer5')) if self.current_data.get('Integer5') else None,
                     "Integer6": int(self.current_data.get('Integer6')) if self.current_data.get('Integer6') else None,
                     "Integer7": int(self.current_data.get('Integer7')) if self.current_data.get('Integer7') else None,
                     "Integer8": int(self.current_data.get('Integer8')) if self.current_data.get('Integer8') else None,
                     "Integer9": int(self.current_data.get('Integer9')) if self.current_data.get('Integer9') else None,
                     "Integer10": int(self.current_data.get('Integer10')) if self.current_data.get(
                         'Integer10') else None,
                     "Integer11": int(self.current_data.get('Integer11')) if self.current_data.get(
                         'Integer11') else None,
                     "Integer12": int(self.current_data.get('Integer12')) if self.current_data.get(
                         'Integer12') else None,
                     "Integer13": int(self.current_data.get('Integer13')) if self.current_data.get(
                         'Integer13') else None,
                     "Integer14": int(self.current_data.get('Integer14')) if self.current_data.get(
                         'Integer14') else None,
                     "Integer15": int(self.current_data.get('Integer15')) if self.current_data.get(
                         'Integer15') else None,
                     "Text1": self.current_data.get('Text1'), "Text2": self.current_data.get('Text2'),
                     "Text3": self.current_data.get('Text3'),
                     "Text4": self.current_data.get('Text4'), "Text5": self.current_data.get('Text5'),
                     "Text6": self.current_data.get('Text6'),
                     "Text7": self.current_data.get('Text7'), "Text8": self.current_data.get('Text8'),
                     "Text9": self.current_data.get('Text9'),
                     "Text10": self.current_data.get('Text10'), "Text11": self.current_data.get('Text11'),
                     "Text12": self.current_data.get('Text12'),
                     "Text13": self.current_data.get('Text13'), "Text14": self.current_data.get('Text14'),
                     "Text15": self.current_data.get('Text15')
                     }
        # print self.data

        self.lambda_function('candidate_duplicate_check')
        self.headers['APP-NAME'] = 'crpo'

        r = requests.post(self.webapi, headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
        print(r.headers)

        time.sleep(1)
        resp_dict = json.loads(r.content)
        self.is_duplicate = resp_dict["IsDuplicate"]
        # print self.is_duplicate
        self.message = resp_dict['Message']

        if self.is_duplicate:
            self.is_duplicate1 = "Duplicate"
            # print self.is_duplicate1
        else:
            self.is_duplicate1 = "NotDuplicate"
            # print self.is_duplicate1

        if self.is_duplicate1 == self.current_data.get('ExpectedOutput'):
            self.style6 = self.style14
        else:
            self.style6 = self.style13

        self.excelwrite(self.message)

    def excelwrite(self, message):

        if message == self.current_data.get('Message'):
            self.status = "Pass"
            style = self.style14
            self.ws.write(self.rowsize, 0, self.status, self.style26)
            self.success_case_01 = 'Pass'
        else:
            self.status = "Fail"
            style = self.style3
            self.ws.write(self.rowsize, 0, self.status, self.style3)

        self.ws.write(self.rowsize, 1, self.current_data.get('CandidateName'), self.style12)
        self.ws.write(self.rowsize, 2, self.current_data.get('FirstName'), self.style12)
        self.ws.write(self.rowsize, 3, self.current_data.get('MiddleName'), self.style12)
        self.ws.write(self.rowsize, 4, self.current_data.get('LastName'), self.style12)
        self.ws.write(self.rowsize, 5, self.current_data.get('EmailAddress'), self.style12)
        self.ws.write(self.rowsize, 6, self.current_data.get('MobileNumber'), self.style12)
        self.ws.write(self.rowsize, 7, self.current_data.get('PhoneNumber'), self.style12)
        self.ws.write(self.rowsize, 8, self.current_data.get('MaritalStatus'), self.style12)
        self.ws.write(self.rowsize, 9, self.current_data.get('Gender'), self.style12)
        self.ws.write(self.rowsize, 10, self.date_of_birth, self.style12)
        self.ws.write(self.rowsize, 11, self.current_data.get('Pancard'), self.style12)
        self.ws.write(self.rowsize, 12, self.current_data.get('Passport'), self.style12)
        self.ws.write(self.rowsize, 13, self.current_data.get('Aadhar'), self.style12)
        self.ws.write(self.rowsize, 14, self.current_data.get('USN'), self.style12)
        self.ws.write(self.rowsize, 15, self.current_data.get('College'), self.style12)
        self.ws.write(self.rowsize, 16, self.current_data.get('Degree'), self.style12)
        self.ws.write(self.rowsize, 17, self.current_data.get('Location'), self.style12)
        self.ws.write(self.rowsize, 18, self.current_data.get('TotalExperienceInMonths'), self.style12)

        self.ws.write(self.rowsize, 19, self.current_data.get('LinkedIn'), self.style12)
        self.ws.write(self.rowsize, 20, self.current_data.get('Facebook'), self.style12)
        self.ws.write(self.rowsize, 21, self.current_data.get('Twitter'), self.style12)

        self.ws.write(self.rowsize, 22, self.current_data.get('Text1'), self.style12)
        self.ws.write(self.rowsize, 23, self.current_data.get('Text2'), self.style12)
        self.ws.write(self.rowsize, 24, self.current_data.get('Text3'), self.style12)
        self.ws.write(self.rowsize, 25, self.current_data.get('Text4'), self.style12)
        self.ws.write(self.rowsize, 26, self.current_data.get('Text5'), self.style12)
        self.ws.write(self.rowsize, 27, self.current_data.get('Text6'), self.style12)
        self.ws.write(self.rowsize, 28, self.current_data.get('Text7'), self.style12)
        self.ws.write(self.rowsize, 29, self.current_data.get('Text8'), self.style12)
        self.ws.write(self.rowsize, 30, self.current_data.get('Text9'), self.style12)
        self.ws.write(self.rowsize, 31, self.current_data.get('Text10'), self.style12)
        self.ws.write(self.rowsize, 32, self.current_data.get('Text11'), self.style12)
        self.ws.write(self.rowsize, 33, self.current_data.get('Text12'), self.style12)
        self.ws.write(self.rowsize, 34, self.current_data.get('Text13'), self.style12)
        self.ws.write(self.rowsize, 35, self.current_data.get('Text14'), self.style12)
        self.ws.write(self.rowsize, 36, self.current_data.get('Text15'), self.style12)

        self.ws.write(self.rowsize, 37, self.current_data.get('Integer1'), self.style12)
        self.ws.write(self.rowsize, 38, self.current_data.get('Integer2'), self.style12)
        self.ws.write(self.rowsize, 39, self.current_data.get('Integer3'), self.style12)
        self.ws.write(self.rowsize, 40, self.current_data.get('Integer4'), self.style12)
        self.ws.write(self.rowsize, 41, self.current_data.get('Integer5'), self.style12)
        self.ws.write(self.rowsize, 42, self.current_data.get('Integer6'), self.style12)
        self.ws.write(self.rowsize, 43, self.current_data.get('Integer7'), self.style12)
        self.ws.write(self.rowsize, 44, self.current_data.get('Integer8'), self.style12)
        self.ws.write(self.rowsize, 45, self.current_data.get('Integer9'), self.style12)
        self.ws.write(self.rowsize, 46, self.current_data.get('Integer10'), self.style12)
        self.ws.write(self.rowsize, 47, self.current_data.get('Integer11'), self.style12)
        self.ws.write(self.rowsize, 48, self.current_data.get('Integer12'), self.style12)
        self.ws.write(self.rowsize, 49, self.current_data.get('Integer13'), self.style12)
        self.ws.write(self.rowsize, 50, self.current_data.get('Integer14'), self.style12)
        self.ws.write(self.rowsize, 51, self.current_data.get('Integer15'), self.style12)
        self.ws.write(self.rowsize, 52, self.current_data.get('DuplicationRuleText'), self.style12)

        self.ws.write(self.rowsize, 53, self.current_data.get('ExpectedOutput'), self.style12)
        self.ws.write(self.rowsize, 54, self.is_duplicate1, self.style14)

        self.ws.write(self.rowsize, 55, self.current_data.get('Message'), self.style12)

        self.ws.write(self.rowsize, 56, message, style)

        self.rowsize = self.rowsize + 1
        self.wb_Result.save(output_paths.outputpaths['Duplication_rule_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

        self.success_case_01 = {}

    def overall_status(self):
        self.ws.write(0, 0, 'Duplication Check', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, self.tot, self.style24)
        ob.wb_Result.save(output_paths.outputpaths['Duplication_rule_Output_sheet'])


ob = VerifyDuplicationRule()
ob.overall_status()
