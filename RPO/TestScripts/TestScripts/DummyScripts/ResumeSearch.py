import ast
import math
import dateutil.parser
import requests
import json
import copy
import mysql
import xlrd
import xlwt
import datetime
from mysql import connector
class Excel_Data:
    def __init__(self):
        self.xl_json_request = []
        self.xl_excepted_candidate_id = []
        self.rownum = 1
        self.boundary_range = [1119500,1265125,1120101,1217454,1120106,1120037,1223268,1217446,1117681,1222591,1116524,
                               1119509,1228469,1218132,1222838,1116199,1117669,1223237,1219415,1223794,1120109,1217637,
                               1119500,1265125,1217454,1120106,1120037,1223268,1217446,1117681,1222591,1119509,1117669,
                               1116524,1116199,1222838,1223237,1218132,1228469,1217637,1219415,1119500,1120101,1117681,
                               1116524,1120109,1265125,1119500,1120101,1217446,1116524,1217454,1116199,1218132,1217637,
                               1116524,1217454,1120037,1228469,1218132,1222838,1217637,1223237,1116569,1264533,1262975,
                               1217539,1217544,1116502,1219552,1217525,1264523,1116410,1217480,1117683,1219465,1119518,
                               1116435,1120130,1217525,1217544,1217539,1217532,1265118,1116457,1116571,1116400,1219467,
                               1222518,1117686,1217618,1120143,1217566,1217608,1116341,1116474,1217627,1116351,1116313,
                               1217742,1116341,1117813]
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Resume Search Result')
        self.ws.write(0, 0, 'Request', self.__style0)
        self.ws.write(0, 1, 'API Count', self.__style0)
        self.ws.write(0, 2, 'Expected Candidate Id\'s', self.__style0)
        self.ws.write(0, 3, 'Not Matched Id\'s', self.__style0)
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "rpo@1234", "TenantAlias": "rpotestone", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=True)
        self.TokenVal = response.json()
        print self.TokenVal.get("Token")
        wb = xlrd.open_workbook('C:\PythonAutomation\InputForJobOfferResumeSearch\ResumeSerachInput.xls')
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_json_request.append(rows[0])
            self.xl_excepted_candidate_id.append(str(rows[1]))
        local = self.xl_excepted_candidate_id
        length = len(self.xl_excepted_candidate_id)
        self.new_local = []
        for i in range(0, length):
            j = [int(float(b)) for b in local[i].split(',')]
            self.new_local.append(j)
        self.xl_expected = self.new_local
    def json_data(self):
        r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_candidates/", headers=self.headers,
                          data=json.dumps(self.data, default=str), verify=False)
        # print self.data
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        print resp_dict
        if self.status == 'OK':
            self.count = resp_dict['TotalItem']
            self.total_pages1 = float(self.count)/200
            self.total_pages = math.ceil(self.total_pages1)
            self.total_pages = int(self.total_pages)
        else:
            self.count = "400000000000000"
    def json_data_iteration(self, data, iter):
        iter += 1
        self.actual_ids = []
        for i in range(1, iter):
            self.data["PagingCriteria"]["PageNo"] = i
            r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_candidates/", headers=self.headers,
                              data=json.dumps(data, default=str), verify=False)
            resp_dict = json.loads(r.content)
            for element in resp_dict["Candidates"]:
                self.actual_ids.append(element["Id"])
    def all(self):
        tot_len = len(self.xl_json_request)
        for i in range(0, tot_len):
            print "Iteration Count :- %s " % i
            self.xl_request= json.loads(self.xl_json_request[i])
            self.xl_request1 = copy.deepcopy(self.xl_request)
            if self.xl_request.get("CandidateId"):
                self.xl_request["CandidateId"] = self.boundary_range
            else:
                val = [("CandidateIds", self.boundary_range)]
                id_filter = dict(val)
                self.xl_request.update(id_filter)
            self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
            self.data = {"PagingCriteria": {"MaxResults": 200, "PageNo": 1}, "CandidateFilters": self.xl_request}
            print self.data
            self.json_data()
            self.total_api_count = self.count
            if self.count != "400000000000000":
                self.data["PagingCriteria"] = {"IsRefresh": False, "MaxResults": 200, "PageNo": 1, "ObjectState": 0}
                print self.data
            # print self.total_pages
            self.json_data_iteration(self.data, self.total_pages)
            self.mismatched_id = set(self.xl_expected[i]) - set(self.actual_ids)
            expected_id = str(self.xl_expected[i])
            expected_id = expected_id.strip('[]')
            mismatched_id = str(list(self.mismatched_id))
            mismatched_id = mismatched_id.strip('[]')
            self.ws.write(self.rownum, 0, str(self.xl_request1))
            if self.total_api_count == self.xl_excepted_candidate_id:
                self.ws.write(self.rownum, 1, self.total_api_count, self.__style3)
                self.ws.write(self.rownum, 2, expected_id, self.__style3)
            elif self.total_api_count == '400000000000000':
                print "API Failed"
                self.ws.write(self.rownum, 1, "API Failed", self.__style2)
                self.ws.write(self.rownum, 2, expected_id, self.__style3)
                self.ws.write(self.rownum, 3, "API Failed", self.__style2)
            else:
                print "this is else part \ n"
                self.ws.write(self.rownum, 1, self.total_api_count, self.__style3)
                self.ws.write(self.rownum, 2, expected_id, self.__style3)
                self.ws.write(self.rownum, 3, mismatched_id, self.__style2)
            self.wb_result.save(
                'C:\PythonAutomation\SearchInResumeResults/'
                + self.__current_DateTime + '_Resume_Search.xls')
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1
print "Resume Search Script Started"
xlob = Excel_Data()
xlob.all()
print "Completed Successfully "