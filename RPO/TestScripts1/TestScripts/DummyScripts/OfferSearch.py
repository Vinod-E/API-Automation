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
        # self.job_filters = ['JobIds','Name','NoOfOpeningsFrom','NoOfOpeningsTo','RoleIds','OwnerIds','LocationIds','IsJobUtilized',
        #                     'DepartmentIds','JobCodeIds','SalaryStart','SalaryEnd','ExperienceStart','ExperienceEnd','IsJobPosted',
        #                     'SkillIds']

        self.xl_json_request = []
        self.xl_excepted_offer_id = []
        self.rownum = 1

        self.boundary_range = [27669,27665,27664,27663,27662,27661,27660,27659,27624,27623,27622,27616,27607,27601,
                               27599,27598,27596,27595,27594,27591,27585,27582,27580,27086,27117,27559,27645,27114,
                               27638,27326,27639,27489,27617,27618,27619,27623,27622,27630,27660,27621,27584,27628,
                               27586,27556,27483]

        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')

        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Offer Search Result')
        self.ws.write(0, 0, 'Request', self.__style0)
        self.ws.write(0, 1, 'API Count', self.__style0)
        self.ws.write(0, 2, 'DB Count', self.__style0)
        self.ws.write(0, 3, 'Expected Offer Id\'s', self.__style0)
        self.ws.write(0, 4, 'Not Matched Id\'s', self.__style0)

        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()
        # self.tenant_id = 1782
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")

        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "rpotestone", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=True)
        self.TokenVal = response.json()
        print self.TokenVal.get("Token")

        wb = xlrd.open_workbook('/home/sanjeev/TestScripts/TestScripts/OfferSearchInput.xls')
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_json_request.append(rows[0])
            self.xl_excepted_offer_id.append(str(rows[1]))

        local = self.xl_excepted_offer_id
        print type(local)
        length = len(self.xl_excepted_offer_id)
        self.new_local = []

        for i in range(0, length):
            j = [int(float(b)) for b in local[i].split(',')]
            self.new_local.append(j)
        self.xl_expected = self.new_local

    def json_data(self):
        r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_offers/", headers=self.headers,
                          data=json.dumps(self.data, default=str), verify=False)
        print self.data
        # print r.content
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        print resp_dict

        if self.status == 'OK':
            self.count = resp_dict['TotalItemCount']
            self.total_pages1 = float(self.count)/200
            self.total_pages = math.ceil(self.total_pages1)
            self.total_pages = int(self.total_pages)

            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count

    def json_data_iteration(self, data, iter):
        iter += 1
        self.actual_ids = []
        for i in range(1, iter):
            self.data["PagingCriteria"]["PageNo"] = i
            r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_offers/", headers=self.headers,
                              data=json.dumps(data, default=str), verify=False)
            resp_dict = json.loads(r.content)
            # print resp_dict
            for element in resp_dict["Offers"]:
                self.actual_ids.append(element["CandidateId"])
                # print element1
                # print len(self.actual_ids)
                # print self.actual_ids

    def all(self):
        tot_len = len(self.xl_json_request)
        for i in range(0, tot_len):
            print "Iteration Count :- %s " % i
            # self.xl_request = ast.literal_eval(self.xl_json_request[i])
            self.xl_request= json.loads(self.xl_json_request[i])
            self.xl_request1 = copy.deepcopy(self.xl_request)

            if self.xl_request.get("CandidateId"):
                self.xl_request["CandidateId"] = self.boundary_range
                # print self.xl_request
            else:
                val = [("CandidateId", self.boundary_range)]
                id_filter = dict(val)
                self.xl_request.update(id_filter)
                # print self.xl_request
            # self.ws.write(self.rownum, 0, str(self.xl_request1))
            # all_keys = self.xl_request.keys()

            self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
            self.data = {"PagingCriteria":{"MaxResults":200,"PageNo":1},"GetAllOffersOption":"2","Filters":self.xl_request}
            print self.data
            self.json_data()
            self.total_api_count = self.count
            if self.count != "400000000000000":
                self.data["PagingCriteria"] = {"IsRefresh": False, "MaxResults": 200, "PageNo": 1, "ObjectState": 0}
                # print self.data
            print self.total_pages
            self.json_data_iteration(self.data, self.total_pages)
            # print self.xl_expected[i]
            # print type(self.xl_expected[i])
            # print self.actual_ids
            # print type(self.actual_ids)
            self.mismatched_id = set(self.xl_expected[i]) - set(self.actual_ids)
            # print self.mismatched_id

            self.Query_Generation()

            expected_id = str(self.xl_expected[i])
            expected_id = expected_id.strip('[]')
            mismatched_id = str(list(self.mismatched_id))
            mismatched_id = mismatched_id.strip('[]')

            self.ws.write(self.rownum, 0, str(self.xl_request1))
            if self.total_api_count == self.Query_Result1:
                self.ws.write(self.rownum, 1, self.total_api_count, self.__style3)
                self.ws.write(self.rownum, 2, self.Query_Result1, self.__style3)
                self.ws.write(self.rownum, 3, expected_id, self.__style3)
                self.ws.write(self.rownum, 4, mismatched_id, self.__style2)

            elif self.total_api_count == '400000000000000':
                print "API Failed"
                self.ws.write(self.rownum, 1, "API Failed", self.__style2)
                self.ws.write(self.rownum, 2, self.Query_Result1, self.__style2)
                self.ws.write(self.rownum, 3, expected_id, self.__style3)
                self.ws.write(self.rownum, 4, "API Failed", self.__style2)
            else:
                print "this is else part \ n"
                self.ws.write(self.rownum, 1, self.total_api_count, self.__style2)
                self.ws.write(self.rownum, 2, self.Query_Result1, self.__style2)
                self.ws.write(self.rownum, 3, expected_id, self.__style3)
                self.ws.write(self.rownum, 4, mismatched_id, self.__style2)
            self.wb_result.save(
                '/home/sanjeev/TestScripts/TestScripts/OfferSearchResults/'
                + self.__current_DateTime + '_Combined_Offer_Search.xls')
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1



    def Query_Generation(self):
        select_str = "select count(distinct(ap.id)) from applicant_statuss ap " \
                     "inner join candidates c on c.id = ap.candidate_id " \
                     "inner join jobs j on j.id = ap.job_id " \
                     "inner join resume_statuss stat on stat.id=ap.current_status_id " \
                     "inner join resume_statuss stag on stag.id=stat.resumestatus_id " \
                     "left join offers o on o.id = ap.offer_id "
        a = self.xl_request.get("CandidateId")
        values = ','.join(str(v) for v in a)
        where_str = "ap.tenant_id=1782 and stag.base_name='Offer'"


        if self.xl_request.get("CandidateName"):
            where_str += " and c.candidate_name like '%{}%' ".format(self.xl_request.get("CandidateName"))

        if self.xl_request.get("JobLocationIds"):
            where_str += " and j.location_id ={} ".format(self.xl_request.get("JobLocationIds")[0])
        if self.xl_request.get("SourceIds"):
            where_str += " and c.original_source_id ={} ".format(self.xl_request.get("SourceIds")[0])
        if self.xl_request.get("DepartmentId"):
            where_str += " and j.department_id ={} ".format(self.xl_request.get("DepartmentId"))
        if self.xl_request.get("JobIds"):
            a = self.xl_request.get("JobIds")
            values = ','.join(str(v) for v in a)
            where_str += " and j.id in(%s) " % values
        if self.xl_request.get("JoiningDateFrom") and self.xl_request.get("JoiningDateTill"):
            joining_date_from = self.xl_request.get("JoiningDateFrom")
            joining_date_till = self.xl_request.get("JoiningDateTill")
            where_str += " and o.date_of_joining between '%s' and '%s' " % (joining_date_from, joining_date_till)
        if self.xl_request.get("ModifiedBy"):
            where_str += " and o.modified_by ={} ".format(self.xl_request.get("ModifiedBy"))
        if self.xl_request.get("OfferCreatedFrom") and self.xl_request.get("OfferCreatedTill"):
            offer_created_from = self.xl_request.get("OfferCreatedFrom")
            offer_created_till = self.xl_request.get("OfferCreatedTill")
            where_str += " and o.created_on between '%s' and '%s' " % (offer_created_from, offer_created_till)
        if self.xl_request.get("StageId"):
            where_str += " and stag.id ={} ".format(self.xl_request.get("StageId"))
        if self.xl_request.get("StatusIds"):
            a = self.xl_request.get("StatusIds")
            values = ','.join(str(v) for v in a)
            where_str += " and stat.id in(%s) " % values

        # if self.xl_request.get("StatusIds"):
        #     where_str += " and stat.id ={} ".format(self.xl_request.get("StatusIds")[0])

        final_qur = ""
        if where_str:
            final_qur = select_str + " where " + where_str
            self.query = final_qur
        if final_qur:
            try:
                self.cursor.execute(final_qur)
                Query_Result = self.cursor.fetchone()
                print Query_Result
                print final_qur
                self.Query_Result1 = Query_Result[0]
            except Exception as e:
                print e



        # def date_converter(self, input_date):
        #     converted_utc_date = dateutil.parser.parse(input_date)
        #     converted_local_date = converted_utc_date.astimezone(dateutil.tz.tzlocal()).replace(tzinfo=None)
        #     date = converted_local_date.strftime("%Y-%m-%d")
        #     return date

print "Offer Search Script Started"
xlob = Excel_Data()
xlob.all()
print "Completed Successfully "
