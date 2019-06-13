import requests
import json
import copy
import mysql
import xlrd
import xlwt
import datetime
from hpro_automation import (login, api, styles, input_paths, output_paths)


class Excel_Data(login.CRPOLogin, styles.FontColor):

    def __init__(self):
        super(Excel_Data, self).__init__()

        self.xl_json_request = []
        self.xl_excepted_job_id = []
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
        self.ws = self.wb_result.add_sheet('Job Search Result')
        self.ws.write(0, 0, 'Request', self.__style0)
        self.ws.write(0, 1, 'API Count', self.__style0)
        self.ws.write(0, 2, 'DB Count', self.__style0)
        self.ws.write(0, 3, 'Expected Job Id\'s', self.__style0)
        self.ws.write(0, 4, 'Not Matched Id\'s', self.__style0)

        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()
        # self.tenant_id = 1782
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")

        wb = xlrd.open_workbook(input_paths.inputpaths['r_job_search_Input_sheet'])
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_json_request.append(rows[0])
            self.xl_excepted_job_id.append(str(rows[1]))

        local = self.xl_excepted_job_id
        print type(local)
        length = len(self.xl_excepted_job_id)
        self.new_local = []

        for i in range(0, length):
            j = [int(float(b)) for b in local[i].split(',')]
            self.new_local.append(j)
        self.xl_expected = self.new_local

    def json_data(self):
        r = requests.post(api.web_api['get_all_jobs'], headers=self.get_token,
                          data=json.dumps(self.data, default=str), verify=False)
        # print r.content
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']

        if self.status == 'OK':
            self.count = resp_dict['TotalItem']
            self.total_pages = int(resp_dict['TotalPages'])
            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count

    def json_data_iteration(self, data, iter):
        iter += 1
        self.actual_ids = []
        for i in range(1, iter):
            self.data["PagingCriteria"]["PageNo"] = i
            r = requests.post(api.web_api['get_all_jobs'], headers=self.get_token,
                              data=json.dumps(data, default=str), verify=False)
            resp_dict = json.loads(r.content)
            # print resp_dict
            for element in resp_dict["Jobs"]:
                self.actual_ids.append(element["Id"])
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

            if self.xl_request.get("JobIds"):
                self.xl_request["JobIds"] = self.boundary_range
                # print self.xl_request
            else:
                val = [("JobIds", self.boundary_range)]
                id_filter = dict(val)
                self.xl_request.update(id_filter)
                # print self.xl_request
            # self.ws.write(self.rownum, 0, str(self.xl_request1))
            # all_keys = self.xl_request.keys()

            self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.get_token}
            self.data = {"PagingCriteria":{"IsRefresh":False,"IsSpecificToUser":False,"MaxResults":20,"ObjectState":0,"PageNo":1,
                                   "PropertyIds":[]},"GetAllJobsOption":3,"JobFilters":self.xl_request}
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
            self.wb_result.save(output_paths.outputpaths['r_Job_search_output_sheet'])
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1

    def Query_Generation(self):
        select_str = "select count(distinct(j.id)) from jobs j " \
                     "left join job_owners jo on j.id = jo.job_id " \
                     "left join job_postings jp on j.id = jp.job_id " \
                     "left join job_required_skills jr on j.id = jr.job_id "
        a = self.xl_request.get("JobIds")
        values = ','.join(str(v) for v in a)
        # where_str += " and j.id in ( %s ) " % values
        where_str = "j.is_deleted=0 and j.is_archived=0 and j.tenant_id=1782 and j.id in ( %s ) "% values

        if self.xl_request.get("Name"):
            where_str += " and j.job_name like '%{}%' ".format(self.xl_request.get("Name"))

        if self.xl_request.get("NoOfOpeningsFrom") and self.xl_request.get("NoOfOpeningsTo"):
            opening_from = self.xl_request.get("NoOfOpeningsFrom")
            opening_to = self.xl_request.get("NoOfOpeningsTo")
            where_str += " and j.no_of_openings between '%s' and '%s' " % (opening_from, opening_to)

        if self.xl_request.get("OwnerIds") and self.xl_request.get("RoleIds"):
            j_owner = self.xl_request.get("OwnerIds")[0]
            print j_owner
            r_owner = self.xl_request.get("RoleIds")[0]
            print r_owner
            where_str += " and jo.user_id=%s and jo.role_id=%s " % (j_owner, r_owner)

        if self.xl_request.get("LocationIds"):
            # print self.xl_request.get("LocationIds")
            print self.xl_request.get("LocationIds")[0]
            where_str += " and j.location_id ={}".format(self.xl_request.get("LocationIds")[0])

        if self.xl_request.get("IsJobUtilized"):
            where_str += " and j.is_utilized ={} ".format(self.xl_request.get("IsJobUtilized"))

        if self.xl_request.get("DepartmentIds"):
            where_str += " and j.department_id = {} and j.sub_department_id={} "\
                .format(self.xl_request.get("DepartmentIds")[0],self.xl_request.get("SubDepartmentIds")[0])

        if self.xl_request.get("JobCodeIds"):
            where_str += " and j.job_code_id ={} ".format(self.xl_request.get("JobCodeIds")[0])

        if self.xl_request.get("SalaryStart") and self.xl_request.get("SalaryEnd"):
            salary_start = self.xl_request.get("SalaryStart")
            salary_end = self.xl_request.get("SalaryEnd")
            where_str += " and j.salary_start >=%s and j.salary_end<=%s " % (salary_start, salary_end)

        if self.xl_request.get("ExperienceStart") and self.xl_request.get("ExperienceEnd"):
            exp_start = self.xl_request.get("ExperienceStart")
            exp_end = self.xl_request.get("ExperienceEnd")
            where_str += " and j.experience_start >=%s and j.experience_end <=%s " % (exp_start, exp_end)

        if self.xl_request.get("IsJobPosted"):
            if self.xl_request.get("IsJobPosted") == 1:
                where_str += " and j.id in (select distinct(job_id) from job_postings)"
            else:
                where_str += " and j.id not in (select distinct(job_id) from job_postings)"
                # where_str += " and jp.job_id ={} ".format(self.xl_request.get("IsJobPosted"))

        if self.xl_request.get("SkillIds"):
            a = self.xl_request.get("SkillIds")
            values = ','.join(str(v) for v in a)
            where_str += " and jr.skill_id in(%s) "%values

        final_qur = ""
        if where_str:
            final_qur = select_str + " where " + where_str
            self.query = final_qur
        if final_qur:
            try:
                print self.query
                self.cursor.execute(final_qur)
                Query_Result = self.cursor.fetchone()
                print final_qur
                self.Query_Result1 = Query_Result[0]
            except Exception as e:
                print e


print "Combined Search Script Started"
xlob = Excel_Data()
xlob.all()
print "Completed Successfully "
