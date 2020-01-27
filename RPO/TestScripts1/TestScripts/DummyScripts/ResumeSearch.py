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
        self.boundary_range = [1223206,1219480,1262031,1223200,1262030,1223221,1219273,1222598,1222699,1262032,
                               1223215,1218499,1116304,1217982,1223209,1261933,1261929,1223211,1261932,1241494,
                               1223216,1223213,1223214,1260226,1223205,1117736,1219471,1117843,1222677,1260255,
                               1237925,1237926,1217706,1217667,1116447,1228307,1228305,1262035,1223219,1223819,
                               1219281,1223203,1261938,1262040,1223204,1237924,1119475,1119474,1220139,1117665,
                               1116545,1260247,1220102,1217478,1220099,1261965,1262037,1218751,1223218,1262033,
                               1217675,1262036,1261937,1116445,1217456,1116479,1120035,1120086,1116382,1116438,
                               1119476,1217619,1223772,1228306,1217447,1217672,1222652,1217936,1220073,1116644,
                               1117747,1219101,1119520,1117441,1117673,1222559,1223233,1223208,1217674,1116409,
                               1119555,1217969,1119500,1219429,1222703,1220152,1223212,1220098,1116514,1116868,
                               1117841,1222698,1223624,1116268,1116520,1218148,1219231,1116416,1120140,1217605,
                               1217723,1219470,1228453,1116293,1116671,1219228,1222514,1223770,1220148,1120095,
                               1117661,1220087,1223739,1223790,1223782,1223536,1116343,1116387,1119564,1217904,
                               1222528,1222668,1220118,1223180,1116385,1116411,1116500,1116539,1117799,1219450,
                               1220166,1220237,1222608,1116526,1116767,1116718,1119493,1220209,1222639,1222640,
                               1223730,1116183,1219412,1219852,1220159,1216686,1223030,1223170,1223596,1241495,
                               1236610,1116306,1116561,1116544,1116678,1116719,1120109,1220282,1223809,1217705,
                               1223926,1116348,1217468,1222527,1223597,1217713,1223982,1119505,1217499,1217983,
                               1222702,1223706,1228477,1220120,1223807,1116401,1217513,1223579,1120084,1223977,
                               1222536,1223537,1228314,1217573,1117458,1117839,1220112,1223829,1220127,1228328,
                               1116576,1220083,1223549,1223733,1223813,1119560,1223539,1230714,1228336,1220081,
                               1220086,1120129,1223817,1219866,1223777,1224088,1116392]
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
        data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "rpotestone", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=True)
        self.TokenVal = response.json()
        print self.TokenVal.get("Token")
        wb = xlrd.open_workbook('/home/sanjeev/TestScripts/TestScripts/ResumeSerachInput.xls')
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
                '/home/sanjeev/TestScripts/TestScripts/ResumeSearchResults/'
                + self.__current_DateTime + '_Resume_Search.xls')
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1
print "Resume Search Script Started"
xlob = Excel_Data()
xlob.all()
print "Completed Successfully "