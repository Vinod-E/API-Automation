import mysql.connector
import requests
import json
from Job_Search_Input import *
import datetime
import xlwt
from collections import OrderedDict

class AMS_DB_Data:
    def __init__(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()
        self.tenant_id = 1782

    def execute_Query(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchone()
        except:
            print("Hi")

    def ams_Query(self,b):

        self.job_id = "select count(1) from jobs where id =%s  and is_archived=0 and tenant_id=%s;"% (xlob.xl_job_id[b],self.tenant_id)
        # print xlob.xl_job_id[b]
        query = self.job_id
        print query
        C4.execute_Query(query)
        self.db_job_id = self.data[0]
        # print "total job id count is %s"%self.db_job_id

        self.job_name = "select count(1) from jobs where job_name like '%s' and is_archived=0 and tenant_id=%s;" %('%'+xlob.xl_job_name[b] + '%',self.tenant_id)
        query = self.job_name
        print query
        C4.execute_Query(query)
        self.db_job_name = self.data[0]
        # print "total job_name count is %s" % self.db_job_name

        self.no_of_opening = "select count(1) from jobs where no_of_openings between %s and %s  and is_archived=0 and tenant_id=%s;"\
                             % (xlob.xl_openings_from[b],xlob.xl_openings_to[b],self.tenant_id)
        query = self.no_of_opening
        print query
        C4.execute_Query(query)
        self.db_no_of_opening = self.data[0]
        # print "Total No Of Opening count is %s" % self.db_no_of_opening

        self.job_owner = "select count(distinct(jo.job_id)) from job_owners jo  left join jobs j on j.id =jo.job_id " \
                         "where jo.user_id=%s and jo.role_id=%s and j.is_archived=0 and j.tenant_id=%s;" \
                             % (xlob.xl_users[b], xlob.xl_roles[b], self.tenant_id)
        query = self.job_owner
        print query
        C4.execute_Query(query)
        self.db_job_owner = self.data[0]
        # print "Total No Of Opening count is %s" % self.db_no_of_opening

        self.location = "select count(1) from jobs where location_id =%s and is_archived=0 and tenant_id=%s;"% (xlob.xl_location[b],self.tenant_id)
        query = self.location
        print query
        C4.execute_Query(query)
        self.db_location = self.data[0]
        # print "total location count is %s" % self.db_location

        self.hiring_type = "select count(1) from jobs where job_type =%s and is_archived=0 and tenant_id=%s;" % (
        xlob.xl_hiring_type[b], self.tenant_id)
        query = self.hiring_type
        print query
        C4.execute_Query(query)
        self.db_hiring_type = self.data[0]
        # print "total hiring type count is %s" % self.db_hiring_type

        self.is_utilized = "select count(1) from jobs where is_utilized =%s and is_archived=0 and tenant_id=%s;" %(xlob.xl_utilized[b],self.tenant_id)
        query = self.is_utilized
        print query
        C4.execute_Query(query)
        self.db_is_utilized = self.data[0]
        # print "total utilized count is %s" % self.db_is_utilized

        self.department = "select count(1) from jobs where department_id = %s and sub_department_id=%s and is_archived=0 and tenant_id= %s;" %\
                          (xlob.xl_department[b],xlob.xl_sub_department[b],self.tenant_id)
        query = self.department
        print query
        C4.execute_Query(query)
        self.db_department = self.data[0]
        # print "total department count is %s" % self.db_department

        self.job_code = "select count(1) from jobs where job_code_id =%s and is_archived=0 and tenant_id=%s;" % (xlob.xl_job_code[b],self.tenant_id)
        query = self.job_code
        print query
        C4.execute_Query(query)
        self.db_job_code = self.data[0]
        # print "total job code count is %s" % self.db_job_code

        self.ctc_range = "select count(1) from jobs where salary_start >=%s and salary_end<=%s and is_archived=0 and tenant_id=%s;" % (xlob.xl_ctc_from[b],xlob.xl_ctc_to[b],self.tenant_id)
        query = self.ctc_range
        print query
        C4.execute_Query(query)
        self.db_ctc_range = self.data[0]
        # print "total ctc range count is %s" % self.db_ctc_range

        self.exp_range = "select count(1) from jobs where experience_start >=%s and experience_end <=%s and is_archived=0 and tenant_id=%s;" %(xlob.xl_experience_from[b], xlob.xl_experience_to[b],self.tenant_id)
        query = self.exp_range
        print query
        C4.execute_Query(query)
        self.db_exp_range = self.data[0]
        # print "total exp range count is %s" % self.db_exp_range


        self.total_job_count = "select count(id) from jobs where tenant_id=%s and is_archived=0" % \
                               (self.tenant_id)
        query = self.total_job_count
        print query
        C4.execute_Query(query)
        self.total_job_count = self.data[0]
        self.job_posted = "select count(distinct(jp.job_id)) from job_postings jp left join jobs j on j.id = jp.job_id where " \
                          "j.tenant_id=%s;" % (self.tenant_id)
        query = self.job_posted
        print query
        C4.execute_Query(query)
        self.db_job_posted_count = self.data[0]
        # self.db_job_posted = self.db_job_posted_count
        # print "total job posted count is %s" % self.db_job_posted
        if xlob.xl_is_posted[b]== 0:
            self.not_posted_count =  self.total_job_count - self.db_job_posted_count
            self.db_job_posted = self.not_posted_count
        else:
            self.db_job_posted = self.db_job_posted_count

        self.job_skill = "select count(distinct(j.id)) from  job_required_skills jr " \
                         "left join jobs j on j.id=jr.job_id where jr.skill_id=%s " \
                         " and is_archived=0 and j.tenant_id=%s;" % \
            (xlob.xl_job_skills[b],self.tenant_id)
        query = self.job_skill
        print query
        C4.execute_Query(query)
        self.db_job_skill = self.data[0]
        # print "total job skill count is %s" % self.db_job_skill




class api_data:

    def json_data(self,):

        r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_jobs/", headers = self.headers,
                          data = json.dumps(self.data, default=str), verify=False)
        # print r.content
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        # print self.status
        # print type(self.status)
        if self.status == 'OK':
            self.count = resp_dict['TotalItem']
            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count


    def api_main(self,b):
        headers1 = {"content-type": "application/json"}
        data = {"LoginName":"admin","Password":"admin@123","TenantAlias":"rpotestone","UserName":"admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=headers1,
                                 data=json.dumps(data), verify=False)
        abc = response.json()
        # print abc.get("Token")
        # print json.loads(response)
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": abc.get("Token")}

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1},"GetAllJobsOption": 3, "JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1",
                     "JobIds":["27557"],"EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["jobRoleId"] = [xlob.xl_job_id[b]]
        # print self.data["JobFilters"]["jobRoleId"]
        self.data["JobFilters"]["JobIds"] = [xlob.xl_job_id[b]]
        # print self.data["JobFilters"]["JobIds"]
        # print self.data
        ob.json_data()
        self.api_job_id = self.count
        # print "Job Id Count is %s " %self.api_job_id


        self.data = {"PagingCriteria":{"IsRefresh":False,"IsSpecificToUser":False,"MaxResults":20,"ObjectState":0,"PageNo":1},
                      "GetAllJobsOption":3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1","Name":"",
                      "EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["Name"] = xlob.xl_job_name[b]
        # print self.data
        ob.json_data()
        self.api_job_name = self.count
        print "Job name Count is %s " %self.api_job_name


        # self.data = {
        #     "PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
        #                        "PageNo": 1},"GetAllJobsOption": 3, "JobFilters":{"hiringType":"-1","isUtilized":"-1",
        #                         "isJobPosted":"-1","jobRoleId":"","JobIds":[""],
        #                         "EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}
        #
        # self.data["JobFilters"]["jobRoleId"] = xlob.xl_job_id[b]
        # self.data["JobFilters"]["JobIds"] = xlob.xl_job_id[b]
        # # print self.data
        # ob.json_data()
        # self.api_job_name = self.count
        # # print "Job Role Id Count is %s " %self.api_job_id

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters": {"hiringType": "-1", "isUtilized": "-1",
                     "isJobPosted": "-1", "NoOfOpeningsFrom":"","NoOfOpeningsTo":"",
                     "EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["NoOfOpeningsFrom"] = xlob.xl_openings_from[b]
        self.data["JobFilters"]["NoOfOpeningsTo"] = xlob.xl_openings_to[b]
        # print self.data
        ob.json_data()
        self.api_no_of_opening = self.count
        # print "No Of Opening is %s " %self.api_no_of_opening

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1},"GetAllJobsOption": 3, "JobFilters": {"hiringType": "-1", "isUtilized": "-1",
                     "isJobPosted": "-1","LocationIds": "", "EventSummary": True, "AppSummary": True,
                               "StatusId": [1919,1920,14989], "IsOpenReq": False}}

        self.data["JobFilters"]["LocationIds"] = [xlob.xl_location[b]]
        # print self.data
        ob.json_data()
        self.api_location = self.count
        # print "Location Count is %s " %self.api_location

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1",
                     "isJobPosted":"-1","OwnerIds":[""],"roleId":"","RoleIds":[],"EventSummary":True,"AppSummary":True,
                     "StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["OwnerIds"] = [xlob.xl_users[b]]
        self.data["JobFilters"]["roleId"] = xlob.xl_roles[b]
        self.data["JobFilters"]["RoleIds"] = [xlob.xl_roles[b]]
        # print self.data
        ob.json_data()
        self.api_owner_id = self.count
        # print "Owner Count is %s " %self.api_owner_id

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"1","isUtilized":"-1","isJobPosted":"-1",
                     "ReqTypeBoolean":False,"EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["ReqTypeBoolean"] = xlob.xl_hiring_type[b]
        # print self.data
        ob.json_data()
        self.api_req_type = self.count
        # print "Req Type Count is %s " %self.api_req_type

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"1","isJobPosted":"-1",
                     "IsJobUtilized":True,"EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["IsJobUtilized"] = xlob.xl_utilized[b]
        # print self.data
        ob.json_data()
        self.api_is_utilized = self.count
        # print "Req Type Count is %s " %self.api_req_type

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1",
                     "isJobPosted":"-1","SubDepartmentIds":[""],"DepartmentIds":[],"EventSummary":True,
                     "AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["DepartmentIds"] = [xlob.xl_department[b]]
        self.data["JobFilters"]["SubDepartmentIds"] = [xlob.xl_sub_department[b]]
        # print self.data
        ob.json_data()
        self.api_dept_id = self.count
        # print "Departments Count is %s " %self.api_dept_id

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1",
                     "JobCodeIds":[""],"EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["JobCodeIds"] = [xlob.xl_job_code[b]]
        # print self.data
        ob.json_data()
        self.api_job_code = self.count
        # print "Job Code Count is %s " %self.api_job_code

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1",
                     "SalaryStart":"","SalaryEnd":"","EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["SalaryStart"] = xlob.xl_ctc_from[b]
        self.data["JobFilters"]["SalaryEnd"] = xlob.xl_ctc_to[b]
        # print self.data
        ob.json_data()
        self.api_ctc_range = self.count
        # print "CTC Range Count is %s " %self.api_ctc_range

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1",
                     "ExperienceStart":"","ExperienceEnd":"","EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["ExperienceStart"] = xlob.xl_experience_from[b]
        self.data["JobFilters"]["ExperienceEnd"] = xlob.xl_experience_to[b]
        # print self.data
        ob.json_data()
        self.api_exp_range = self.count
        # print "EXP Range Count is %s " %self.api_exp_range

        self.data = {"PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 20, "ObjectState": 0,
                     "PageNo": 1}, "GetAllJobsOption": 3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"1",
                     "IsJobPosted":True,"EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["IsJobPosted"] = xlob.xl_is_posted[b]
        # print self.data
        ob.json_data()
        self.api_job_posted = self.count
        # print "Job Posted Count is %s " %self.api_job_posted

        self.data = {"PagingCriteria":{"IsRefresh":False,"IsSpecificToUser":False,"MaxResults":20,"ObjectState":0,"PageNo":1},
                     "GetAllJobsOption":3,"JobFilters":{"hiringType":"-1","isUtilized":"-1","isJobPosted":"-1","SkillIds":[""],
                    "EventSummary":True,"AppSummary":True,"StatusId":[1919,1920,14989],"IsOpenReq":False}}

        self.data["JobFilters"]["SkillIds"] = [xlob.xl_job_skills[b]]
        # print self.data["JobFilters"]
        # print self.data
        ob.json_data()
        self.api_job_skill = self.count
        # print "Job Skill Count is %s " %self.api_job_skill





    def Compare_Values(self, search_data,api_value, db_value,mess):
        self.ws.write(self.rowsize, self.a, search_data, self.__style3)
        if (api_value == db_value):
            self.ws.write(self.rowsize+1, self.a, api_value, self.__style1)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style1)
            self.a = self.a + 1
            # print self.a
        elif (api_value == '400000000000000'):
            self.ws.write(self.rowsize+1, self.a, 'API is Throwing Error', self.__style2)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style2)
            self.a = self.a + 1
            # print self.a
        else:
            self.ws.write(self.rowsize+1, self.a, api_value, self.__style2)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style2)
            self.a = self.a + 1
            print self.mess
            print "api Value is %s"% api_value
            # print type(api_value)
            print "DB Value is %s" % db_value
            # print type(api_value)

    def validation_MS_and_UI1(self):
        now = datetime.datetime.now()
        self.rowsize = 1
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        # print self.__current_DateTime
        # print type(self.__current_DateTime)
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Job_Search')
        self.ws.write(0, 0, 'Type', self.__style0)

        self.ws.write(0, 1, 'Job Id', self.__style0)
        self.ws.write(0, 2, 'Job Name', self.__style0)
        self.ws.write(0, 3, 'No Of Openings', self.__style0)
        self.ws.write(0, 4, 'Owners', self.__style0)
        self.ws.write(0, 5, 'Location', self.__style0)
        self.ws.write(0, 6, 'Hiring Type', self.__style0)
        self.ws.write(0, 7, 'Utilized', self.__style0)
        self.ws.write(0, 8, 'Department', self.__style0)
        self.ws.write(0, 9, 'Job Code', self.__style0)
        self.ws.write(0, 10, 'CTC Range', self.__style0)
        self.ws.write(0, 11, 'Exp Range', self.__style0)
        self.ws.write(0, 12, 'Job Posted', self.__style0)
        self.ws.write(0, 13, 'Job Skill', self.__style0)




    def validation_MS_and_UI2(self, b):
        self.a = 1
        # print self.rowsize
        # print self.a
        self.ws.write(self.rowsize, 0, 'Search Keyword', self.__style0)
        self.ws.write(self.rowsize + 1, 0, 'API Count', self.__style0)
        self.ws.write(self.rowsize + 2, 0, 'DB Count', self.__style0)
        # print "Compare method is started"

        self.mess = 'Job Id is  not matched'
        self.Compare_Values(xlob.xl_job_id[b], ob.api_job_id, C4.db_job_id, self.mess)
        self.mess = 'Job Name is  not matched'
        self.Compare_Values(xlob.xl_job_name[b],ob.api_job_name, C4.db_job_name,self.mess)
        self.mess = 'No Of Opening is  not matched'
        xlob.xl_openings = str(xlob.xl_openings_from[b])+ "-" + str(xlob.xl_openings_to[b])
        self.Compare_Values( xlob.xl_openings,ob.api_no_of_opening, C4.db_no_of_opening, self.mess)
        # self.a = self.a + 1
        self.mess = 'Owners is not matched'
        xlob.xl_owners = str(xlob.xl_roles[b]) + " - " + str(xlob.xl_users[b])
        self.Compare_Values(xlob.xl_owners,ob.api_owner_id, C4.db_job_owner, self.mess)
        self.mess = 'Location is not matched'
        self.Compare_Values(xlob.xl_location[b], ob.api_location, C4.db_location, self.mess)
        self.mess = 'Hiring Type is not matched'
        self.Compare_Values(xlob.xl_hiring_type[b], ob.api_req_type, C4.db_hiring_type, self.mess)
        self.mess = 'Is Utilized is not matched'
        self.Compare_Values(xlob.xl_utilized[b], ob.api_is_utilized, C4.db_is_utilized, self.mess)
        self.mess = 'Department is not matched'
        xlob.xl_dept = str(xlob.xl_department[b]) + " - " + str(xlob.xl_sub_department[b])
        self.Compare_Values(xlob.xl_dept, ob.api_dept_id, C4.db_department, self.mess)
        self.mess = 'Job Code is not matched'
        self.Compare_Values(xlob.xl_job_code[b], ob.api_job_code, C4.db_job_code, self.mess)
        self.mess = 'CTC Range is not matched'
        xlob.xl_ctc_rane = str(xlob.xl_ctc_from) + " - " + str(xlob.xl_ctc_to[b])
        self.Compare_Values(xlob.xl_ctc_rane[b], ob.api_ctc_range, C4.db_ctc_range, self.mess)
        self.mess = 'Exp Range is not matched'
        xlob.xl_exp_rane = str(xlob.xl_experience_from[b]) + " - " + str(xlob.xl_experience_to[b])
        self.Compare_Values(xlob.xl_exp_rane, ob.api_exp_range, C4.db_exp_range, self.mess)
        self.mess = 'Job Posted is not matched'
        self.Compare_Values(xlob.xl_is_posted[b], ob.api_job_posted, C4.db_job_posted, self.mess)
        self.mess = 'Job Skill is not matched'
        self.Compare_Values(xlob.xl_job_skills[b], ob.api_job_skill, C4.db_job_skill, self.mess)

        self.wb_Result.save(
                    '/home/sanjeev/TestScripts/TestScripts/JobSearchResults/API_DB_Search_Verification(' + self.__current_DateTime + ').xls')
        self.rowsize = self.rowsize +4

C4 = AMS_DB_Data()
ob = api_data()
print len(xlob.xl_job_id)
tot_count = len(xlob.xl_job_id)
ob.validation_MS_and_UI1()

for b in range(0,tot_count):
    print b
    print "Job Search Script Started....."
    C4.ams_Query(b)
    print "thi is api main"
    ob.api_main(b)
    ob.validation_MS_and_UI2(b)