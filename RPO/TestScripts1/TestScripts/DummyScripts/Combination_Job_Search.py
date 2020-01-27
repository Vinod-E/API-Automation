import unittest
import requests
import json
import mysql
import xlrd
import xlwt
import datetime
from mysql  import connector
from itertools import combinations

class CombinationJobSearch(unittest.TestCase):
    def test_CombinationJobSearch(self):



        self.__style0 = xlwt.easyxf('font: Name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: Name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: Name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: Name Times New Roman, color-index green, bold on')


        self.conn = mysql.connector.connect(host = '35.154.36.218',
                                            database = 'appserver_core',
                                            user = 'hireprouser',
                                            password = 'tech@123')

        self.cursor = self.conn.cursor()
        self.tenant_id = 1782
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d-%m-%y")


        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "rpotestone", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                 headers = header, data = json.dumps(data), verify = True)
        self.TokenVal = response.json()
        print self.TokenVal.get("Token")


        wb = xlrd.open_workbook("/home/sanjeev/TestScripts/TestScripts/Job_Search.xls")
        wb_result = xlwt.Workbook()
        ws = wb_result.add_sheet("Job Saerch Results")
        ws.write(0, 0, 'Search Criteria', self.__style0)
        ws.write(0, 1, 'API Count', self.__style0)
        ws.write(0, 2, 'DB Count', self.__style0)
        sheetName = wb.sheet_names()
        print(sheetName)
        sh1 = wb.sheet_by_index(0)
        n = 1
        rownum = n
        while n < sh1.nrows:
            rows = sh1.row_values(rownum)
            ids = str(rows[0]).split(',')
            Id = {'jobRoleId': ids}
            Name = {'Name': rows[1]}
            #Openings_From = {'NoOfOpeningsFrom': rows[2]}
            #Openings_To = {'NoOfOpeningsTo': rows[3]}
            # Role = {'roleId': rows[4]}
            # Owners = {'OwnerIds': rows[5]}
            Location = {'LocationIds': rows[6]}
            #Hiringtype = {'ReqType': rows[5]}
            #Utilized = {'JobUtilized': rows[6]}
            # Departments = {'DepartmentIds': rows[7]}
            # Sub_Dept = {'SubDepartmentIds': rows[8]}
            Jobcode = {'JobCodeIds': rows[9]}
            # Ctc_From = {'SalaryStart': rows[10]}
            # Ctc_To = {'SalaryEnd': rows[11]}
            # Exp_From = {'ExperienceStart': rows[12]}
            # Exp_To = {'ExperienceEnd    ': rows[13]}
            #Jobposted = {'JobPosted': rows[11]}
            Skill = {'SkillIds': rows[14]}


            array = [Id, Name, Location, Jobcode, Skill]
            a = 1
            for i in array:
                comb = combinations(array, a)
                a = a + 1
                for j in comb:
                    self.b = {}
                    for k in j:
                        self.b.update(k)


                        header = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
                        self.data = {"PagingCriteria":{"IsRefresh":False,"IsSpecificToUser":False,"MaxResults":20,"ObjectState":0},
                                "GetAllJobsOption":3,"JobFilters":self.b}
                        r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_jobs/",
                                          headers = header, data = json.dump(self.data, default=str) , varify = True)

                        print r.content
                        statusCode = json.loads(r.content)['TotalItem']
                        print statusCode
                        x = str(self.b)
                        print x
                        self.Query_Generation()



                        ws.write(rownum, 0, x)
                        if statusCode == self.Query_Result1:
                            ws.write(rownum, 1, statusCode, self.__style3)
                            ws.write(rownum, 2, self.Query_Result1, self.__style3)
                        else:
                            ws.write(rownum, 1, statusCode, self.__style2)
                            ws.write(rownum, 2, self.Query_Result1, self.__style2)
                        wb.result.save('/home/sanjeev/TestScripts/TestScripts/Combination_Job_Search_Results'
                                       + __current_DateTime + 'Job_Search_Xls')
                        rownum = rownum  + 1
            n = n + 1






    def Query_Generation(self):
        select_str = "select count(id) from jobs "
        where_str = ""
        if self.b.get("Name"):
            where_str += "and job_name like '%{}%' ".format(self.b.get("Name"))
        if self.b.get("Location"):
            where_str += "and location_id = %s ".format(self.b.get("Location"))
        if self.b.get("Jobcode"):
            where_str += "and jobcode_id = %s ".format(self.b.get("Jobcode"))
        if self.b.get("Skill"):
            where_str += "and skill_id = %s ".format(self.b.get("Skill"))
        final_qur = ""
        if where_str:
            final_qur = select_str + "where" + where_str.lstrip('and')
        print final_qur
        if final_qur:
            self.cursor.execute(final_qur)
            Query_Result = self.cursor.fetchone()
            self.Query_Result1 = Query_Result[0]
            print self.Query_Result1


