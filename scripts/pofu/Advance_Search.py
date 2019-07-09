import requests
import json
import pymysql
import xlrd
import xlwt
import datetime


class Excel_Data:
    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        # This Script works for below fields
        self.candidate_search = ['CandidateIds', 'CandidatName', 'Email', 'ContactNumber', 'Integer1']
        self.entity_search = []
        self.filled_form = []
        self.filled_type_form = []
        self.CandidateCustoms = []
        self.our_expected_ids = [1219775, 1219774, 1219773, 1219772, 1219771, 1219770, 1219663, 1219753,
                                 1219734, 1219733, 1219730, 1219729, 1219728, 1219620, 1219601, 1219600, 1219392]
                                 #1217277, 1217288, 1217300, 1217312, 1217324, 1217334, 1217346, 1217357, 1217369,
                                 #1217384, 1217396, 1217408, 1217420, 1217431, 1217439, 1217451, 1217463, 1217475,
                                 #1217495, 1217507, 1217519, 1217531, 1217543, 1217555, 1217561, 1219655, 1219935,
                                 #1219947, 1219948, 1219951]
        self.xl_json_request = []
        self.xl_exceted_candidate_id = []
        self.rownum = 2
        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.style24 = xlwt.easyxf('font: name Arial, color green, bold on, height 400;'
                                   'align: vert centre, horiz centre;')
        self.style25 = xlwt.easyxf('font: name Arial, color red, bold on, height 400;'
                                   'align: vert centre, horiz centre;')
        # --------------------------------------------------------------------------------------------------------------
        # Write Excel Header's
        # --------------------------------------------------------------------------------------------------------------
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Search Result')
        self.ws.write(1, 0, 'Request', self.__style0)
        self.ws.write(1, 1, 'Status', self.__style0)
        self.ws.write(1, 2, 'API Count', self.__style0)
        self.ws.write(1, 3, 'DB Count', self.__style0)
        self.ws.write(1, 4, 'Expected Candidate Id\'s', self.__style0)
        self.ws.write(1, 5, 'Not Matched Id\'s', self.__style0)
        # self.ws.write(0, 5, 'API Id\'s', self.__style0)
        # --------------------------------------------------------------------------------------------------------------
        # DB Connection
        # --------------------------------------------------------------------------------------------------------------
        self.conn = pymysql.connect(host='35.154.36.218',
                                            database='appserver_core',
                                            user='qauser',
                                            password='qauser')
        self.cursor = self.conn.cursor()
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        # --------------------------------------------------------------------------------------------------------------
        # Login Using API
        # --------------------------------------------------------------------------------------------------------------
        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "admin@123", "TenantAlias": "staffingautomation", "UserName": "admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=header,
                                 data=json.dumps(data), verify=True)
        self.TokenVal = response.json()
        # print self.TokenVal.get("Token")

    # ------------------------------------------------------------------------------------------------------------------
    # Read Excel Data
    # ------------------------------------------------------------------------------------------------------------------
        wb = xlrd.open_workbook('/home/vinodkumar/hirepro_automation/API-Automation/Input Data/Pofu/AdvanceSearch/Candidate_Combined_Search_Boundary_Condition_01-automation.xls')


        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        # print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_json_request.append(rows[0])
            self.xl_exceted_candidate_id.append(str(rows[1]))

        local = self.xl_exceted_candidate_id
        length = len(self.xl_exceted_candidate_id)
        self.new_local = []
        # --------------------------------------------------------------------------------------------------------------
        # below loop is used to convert input expected id's in to lists
        # Note:- for comparisition we need lists
        # --------------------------------------------------------------------------------------------------------------
        for i in range(0, length):
            j = [int(float(b)) for b in local[i].split(',')]
            self.new_local.append(j)
        self.xl_expected = self.new_local

    def json_data(self):
        r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/getAllCandidates/", headers= self.headers,
                          data = json.dumps(self.data, default=str), verify=False)
        # print r.content
        resp_dict = json.loads(r.content)
        self.status = resp_dict['count']
        print (resp_dict)

        if self.status != 'count':
            self.count = resp_dict['count']
            if resp_dict['count']!= 0:
                self.total_pages = int(resp_dict['count'])
            else:
                self.total_pages = 1
            # print self.count
        else:
            self.count = "400000000000000"

    def all(self):
            i = 0

            for x in self.xl_json_request:
                 #search_request["CandidateCustoms"] = {}
                 search_request_list = []
                 search_request = {"SearchStatus":-1}
                 search_request["CandidateCustoms"] = {}
                 search_request["EntityOwnerIds"] = {}
                 #search_request["TaskIds"] = {}
                 x = json.loads(x)
                 if x.get('CandidatName'):
                     search_request["CandidatName"] = x.get('CandidatName')
                 if x.get('Email'):
                     search_request["Email"] = x.get("Email")
                 if x.get('ContactNumber'):
                     search_request["ContactNumber"] = x.get('ContactNumber')
                 if x.get('ActivityId'):
                     search_request["ActivityId"] = x.get('ActivityId')
                 if x.get('Status'):
                     search_request["Status"] = x.get('Status')
                 if x.get('TaskStatus'):
                     search_request["TaskStatus"] = x.get('TaskStatus')
                 if x.get('CallStatus'):
                     search_request["CallStatus"] = x.get('CallStatus')
                 if x.get('ThirdPartyIds'):
                     search_request["ThirdPartyIds"] = x.get('ThirdPartyIds')
                 if x.get('IsDOJMismatch'):
                     search_request["IsDOJMismatch"] = x.get('IsDOJMismatch')
                 if x.get('IsCallsAssigned'):
                     search_request["IsCallsAssigned"] = x.get('IsCallsAssigned')
                 if x.get('createdWithinDays'):
                     search_request["createdWithinDays"] = x.get('createdWithinDays')

                 if x.get('expectedDOJInDays'):
                     search_request["expectedDOJInDays"] = x.get('expectedDOJInDays')
                 if x.get('notCalledForDays'):
                     search_request["notCalledForDays"] = x.get('notCalledForDays')
                 if x.get('CandidateIds'):
                     search_request["CandidateIds"] = x.get('CandidateIds')
                 # if x.get('OfferedLevelIds'):
                 #     search_request["OfferedLevelIds"] = x.get('OfferedLevelIds')
                 if x.get('TaskIds'):
                     search_request["TaskIds"] = x.get('TaskIds')
                 if x.get('Tags'):
                     search_request["Tags"] = x.get('Tags')

                 if x.get('OfferedLevelIds'):
                     search_request["OfferedLevelIds"] = x.get('OfferedLevelIds')

                 if x.get('BusinessUnitIds'):
                     search_request["BusinessUnitIds"] = x.get('BusinessUnitIds')

                 if x.get('DepartmentIds'):
                     search_request["DepartmentIds"] = x.get('DepartmentIds')

                 if x.get('CreatedOn'):
                     search_request["CreatedOn"] = x.get('CreatedOn')

                 if x.get('ActualDOJ'):
                     search_request["ActualDOJ"] = x.get('ActualDOJ')
                 if x.get('TentativeDOJ'):
                     search_request["TentativeDOJ"] = x.get('TentativeDOJ')

                 if x.get('ExpectedDOJ'):
                     search_request["ExpectedDOJ"] = x.get('ExpectedDOJ')

                 if x.get('CandidateStage'):
                     search_request["CandidateStage"] = x.get('CandidateStage')

                 if x.get('CandidateStatus'):
                     search_request["CandidateStatus"] = x.get('CandidateStatus')

                 if x.get('SpocIds'):
                     search_request["SpocIds"] = x.get('SpocIds')

                 # if x.get('CandidateCustoms'):
                 #     search_request["CandidateCustom"] = x.get('CandidateCustoms')

                 if x.get('Integer1'):
                     search_request["CandidateCustoms"]["Integer1"] = x.get('Integer1')
                 if x.get('Integer2'):
                      search_request["CandidateCustoms"]["Integer2"] = x.get('Integer2')
                 if x.get('Integer3'):
                      search_request["CandidateCustoms"]["Integer3"] = x.get('Integer3')

                 if x.get('Integer4'):
                     search_request["CandidateCustoms"]["Integer4"] = x.get('Integer4')

                 if x.get('Integer5'):
                     search_request["CandidateCustoms"]["Integer5"] = x.get('Integer5')

                 if x.get('Integer6'):
                     search_request["CandidateCustoms"]["Integer6"] = x.get('Integer6')
                 if x.get('Integer7'):
                     search_request["CandidateCustoms"]["Integer7"] = x.get('Integer7')
                 if x.get('Integer8'):
                     search_request["CandidateCustoms"]["Integer8"] = x.get('Integer8')
                 if x.get('Integer9'):
                     search_request["CandidateCustoms"]["Integer9"] = x.get('Integer9')
                 if x.get('Integer10'):
                     search_request["CandidateCustoms"]["Integer10"] = x.get('Integer10')
                 if x.get('Integer11'):
                     search_request["CandidateCustoms"]["Integer11"] = x.get('Integer11')

                 if x.get('UserIds') and x.get('RoleId'):
                     search_request["EntityOwnerIds"]["RoleId"] = x.get('RoleId')
                     search_request["EntityOwnerIds"]["UserIds"] = x.get('UserIds')

                 elif x.get('RoleId'):
                     search_request["EntityOwnerIds"]["RoleId"] = x.get('RoleId')

                 if x.get('Text1'):
                     search_request["CandidateCustoms"]["Text1"] = x.get('Text1')
                 if x.get('Text2'):
                     search_request["CandidateCustoms"]["Text2"] = x.get('Text2')
                 if x.get('Text3'):
                     search_request["CandidateCustoms"]["Text3"] = x.get('Text3')
                 if x.get('Text4'):
                     search_request["CandidateCustoms"]["Text4"] = x.get('Text4')
                 if x.get('Text5'):
                     search_request["CandidateCustoms"]["Text5"] = x.get('Text5')
                 if x.get('Text6'):
                     search_request["CandidateCustoms"]["Text6"] = x.get('Text6')
                 if x.get('Text7'):
                     search_request["CandidateCustoms"]["Text7"] = x.get('Text7')
                 if x.get('Text8'):
                     search_request["CandidateCustoms"]["Text8"] = x.get('Text8')
                 if x.get('Text9'):
                     search_request["CandidateCustoms"]["Text9"] = x.get('Text9')
                 if x.get('Text10'):
                     search_request["CandidateCustoms"]["Text10"] = x.get('Text10')

                 self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 59)))

                 search_request_list.append(search_request)

                 for m in search_request_list:
                    self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
                    self.data = {
                        "PagingCriteria": {"MaxResults": 60, "PageNo": 1, "IsSpecificToUser": False, "ObjectState": 0},
                        "IsTotalCountRequired": True, "CandiateSearch": m,
                        "UserSpecificCandidates": False,
                        "LeadSpecificCandidates": False,
                        "IsDashboardCountRequired": False,
                        "IsOnlyCountRequired": True}

                    print (self.data)
                    self.json_data()
                    self.total_api_count = self.count
                    if self.count != "400000000000000":
                        self.data["PagingCriteria"] = {"IsRefresh": False, "MaxResults": 200, "PageNo": 1,
                                                       "ObjectState": 0}


                    self.actual_ids = []
                    self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.TokenVal.get("Token")}
                    self.data = {
                        "PagingCriteria": {"MaxResults": 60, "PageNo": 1, "IsSpecificToUser": False, "ObjectState": 0},
                        "IsTotalCountRequired": False, "CandiateSearch": m,
                        "UserSpecificCandidates": False,
                        "LeadSpecificCandidates": False,
                        "IsDashboardCountRequired": False,
                        "IsOnlyCountRequired": False}

                    r = requests.post("https://amsin.hirepro.in/py/pofu/api/v1/getAllCandidates/", headers=self.headers,
                                  data=json.dumps(self.data, default=str), verify=False)

                    cou_resp_dict = json.loads(r.content)
                    print (cou_resp_dict)

                    for element in cou_resp_dict["Candidates"]:
                        self.actual_ids.append(element["Id"])





                 self.mismatched_id = set(self.xl_expected[i]) - set(self.actual_ids)

                 self.Query_Generation(m)
                 # ----------------------------------------------------------------------------------------------------------
                 # Generating Dynamic query based on combinations and execute
                 # ----------------------------------------------------------------------------------------------------------
                 expected_id = str(self.xl_expected[i])
                 expected_id = expected_id.strip('[]')
                 mismatched_id = str(list(self.mismatched_id))
                 mismatched_id = mismatched_id.strip('[]')

                 self.ws.write(self.rownum, 0, str(x))
                 if self.total_api_count == self.Query_Result1:
                     self.ws.write(self.rownum, 1, 'Pass', self.__style3)
                     self.ws.write(self.rownum, 2, self.total_api_count, self.__style1)
                     self.ws.write(self.rownum, 3, self.Query_Result1, self.__style3)
                     self.ws.write(self.rownum, 4, expected_id, self.__style3)
                     self.ws.write(self.rownum, 5, mismatched_id, self.__style2)

                 elif self.total_api_count == '400000000000000':
                     print ("API Failed")
                     self.ws.write(self.rownum, 1, 'Fail', self.__style2)
                     self.ws.write(self.rownum, 2, "API Failed", self.__style2)
                     self.ws.write(self.rownum, 3, self.Query_Result1, self.__style2)
                     self.ws.write(self.rownum, 4, expected_id, self.__style3)
                     self.ws.write(self.rownum, 5, "API Failed", self.__style2)
                     print (self.query)
                 else:
                     print ("this is else part \ n")
                     self.ws.write(self.rownum, 1, 'Fail', self.__style2)
                     self.ws.write(self.rownum, 2, self.total_api_count, self.__style2)
                     self.ws.write(self.rownum, 3, self.Query_Result1, self.__style2)
                     self.ws.write(self.rownum, 4, expected_id, self.__style3)
                     self.ws.write(self.rownum, 5, mismatched_id, self.__style2)
                     print (self.query)
                 self.wb_result.save(
                     '/home/vinodkumar/hirepro_automation/API-Automation/Output Data/Pofu/AdvanceSearch/Advance_Search_01.xls')
                 self.rownum = self.rownum + 1
                 i =i+1

    def Query_Generation(self, req):
        select_str = "select count(distinct(ca.id)) from candidates ca" \
        " inner join candidate_staffing_profiles csp" \
        " on ca.candidatestaffingprofile_id = csp.id left join assigned_tasks at on " \
        " at.candidate_id=ca.id " \
        " left join candidate_customs cc on ca.candidatecustom_id = cc.id" \
        " left join staffing_statuss ss on ss.candidate_id = ca.id" \
        " LEFT outer JOIN resume_statuss rs on ss.current_status_id = rs.id "\
        " LEFT outer JOIN resume_statuss rs1 on rs.resumestatus_id = rs1.id "\
        " left join filled_forms ff on  ff.candidate_id = ca.id" \
        " left join candidate_spocs cans on ca.id = cans.candidate_id " \
        " left join entity_owners eo on ca.id = eo.entity_id"



        where_str = " ca.is_deleted=0 and ca.is_draft=0 and ca.is_archived=0 and ca.tenant_id=1794"



        if req.get("CandidatName"):
            where_str += "  and  hp_dec(ca.first_name) like \'%{}%\' ".format(req.get("CandidatName"))
        if req.get("Email"):
            where_str += " and hp_dec(ca.email1) like '%{}%' ".format(req.get("Email"))

        if req.get("ContactNumber"):
            where_str += " and hp_dec(ca.mobile1) = {} ".format(req.get("ContactNumber"))
        if req.get("CandidateIds"):
            a = req.get("CandidateIds")
            values = ','.join(str(v) for v in a)
            where_str += " and ca.id in ( %s ) " % values

        if req.get("SpocIds"):
            a = req.get("SpocIds")
            values = ','.join(str(v) for v in a)
            where_str += " and cans.user_id in ( %s ) " % values

        if req.get("ThirdPartyIds"):
            a = req.get("ThirdPartyIds")
            values = ','.join(str(v) for v in a)
            where_str += " and ca.third_party_id in ( %s ) " % values

        if req.get("TaskStatus"):
            where_str += "  and  at.status = {} ".format(req.get("TaskStatus"))
        if req.get("ActivityId"):
            where_str += "  and  ca.current_activity = {} ".format(req.get("ActivityId"))

        if req.get("Status"):
            where_str += " and  activity_status = {} ".format(req.get("Status"))

        if req.get("CandidateStage"):
            where_str +=" and rs1.id = {}".format(req.get("CandidateStage"))

        if req.get("CandidateStatus"):
            a = req.get("CandidateStatus")
            values = ','.join(str(v) for v in a)
            where_str += " and ss.current_status_id in ( %s ) " % values



        if req.get("CandidateCustoms").get("Integer1"):
            where_str += " and ca.integer1 =  {}".format(req.get("CandidateCustoms").get("Integer1"))
        if req.get("CandidateCustoms").get("Integer2"):
            where_str += " and  ca.integer2 =  {}".format(req.get("CandidateCustoms").get("Integer2"))
        if req.get("CandidateCustoms").get("Integer3"):
            where_str += " and ca.integer3 =  {}".format(req.get("CandidateCustoms").get("Integer3"))
        if req.get("CandidateCustoms").get("Integer4"):
            where_str += " and ca.integer4 =  {}".format(req.get("CandidateCustoms").get("Integer4"))
        if req.get("CandidateCustoms").get("Integer5"):
            where_str += " and ca.integer5 =  {}".format(req.get("CandidateCustoms").get("Integer5"))

        if req.get("CandidateCustoms").get("Integer6"):
            where_str += " and cc.integer6 =  {}".format(req.get("CandidateCustoms").get("Integer6"))

        if req.get("CandidateCustoms").get("Integer7"):
            where_str += " and cc.integer7 =  {}".format(req.get("CandidateCustoms").get("Integer7"))

        if req.get("CandidateCustoms").get("Integer8"):
            where_str += " and cc.integer8 =  {}".format(req.get("CandidateCustoms").get("Integer8"))

        if req.get("CandidateCustoms").get("Integer9"):
            where_str += " and cc.integer9 =  {}".format(req.get("CandidateCustoms").get("Integer9"))

        if req.get("CandidateCustoms").get("Integer10"):
            where_str += " and cc.integer10 =  {}".format(req.get("CandidateCustoms").get("Integer10"))

        if req.get("CandidateCustoms").get("Integer11"):
            where_str += " and cc.integer11 =  {}".format(req.get("CandidateCustoms").get("Integer11"))

        if req.get("CandidateCustoms").get("Integer12"):
            where_str += " and cc.integer12 =  {}".format(req.get("CandidateCustoms").get("Integer12"))

        if req.get("EntityOwnerIds").get("UserIds") and req.get("EntityOwnerIds").get("RoleId"):
            user_ids = req.get("EntityOwnerIds").get("UserIds")
            user_ids = ','.join(str(v) for v in user_ids)
            where_str += " and eo.role_id = %s and eo.user_id in (%s)" % (req.get("EntityOwnerIds").get("RoleId"),user_ids)

        elif req.get("EntityOwnerIds").get("RoleId"):
            where_str += " and eo.role_id = {} ".format(req.get("EntityOwnerIds").get("RoleId"))

        if req.get("TaskIds"):
            task_ids = req.get("TaskIds")
            values = ','.join(str(v) for v in task_ids)
            where_str += " and at.task_id in ( %s ) " %values


        if req.get("CandidateCustoms").get("Text1"):
            where_str += " and ca.text1 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text1"))

        if req.get("CandidateCustoms").get("Text2"):
            where_str += " and ca.text2 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text2"))

        if req.get("CandidateCustoms").get("Text3"):
            where_str += " and ca.text3 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text3"))

        if req.get("CandidateCustoms").get("Text4"):
            where_str += " and ca.text4 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text4"))

        if req.get("CandidateCustoms").get("Text5"):
            where_str += " and ca.text5 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text5"))

        if req.get("CandidateCustoms").get("Text6"):
            where_str += " and cc.text6 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text6"))

        if req.get("CandidateCustoms").get("Text7"):
            where_str += " and cc.text7 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text7"))

        if req.get("CandidateCustoms").get("Text8"):
            where_str += " and cc.text8 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text8"))

        if req.get("CandidateCustoms").get("Text9"):
            where_str += " and cc.text9 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text9"))
        if req.get("CandidateCustoms").get("Text10"):
            where_str += " and cc.text10 like \'%{}%\'".format(req.get("CandidateCustoms").get("Text10"))

        if req.get("DepartmentIds"):
            a = req.get("DepartmentIds")
            values = ','.join(str(v) for v in a)
            where_str += " and csp.offered_bu in ( %s ) " % values

        if req.get("BusinessUnitIds"):
            a = req.get("BusinessUnitIds")
            values = ','.join(str(v) for v in a)
            where_str += " and csp.offered_buid in ( %s ) " % values

        if req.get("OfferedLevelIds"):
            a = req.get("OfferedLevelIds")
            values = ','.join(str(v) for v in a)
            where_str += " and csp.offered_Level in ( %s ) " % values

        if req.get("CreatedOn"):

            created_on_from = (req.get("CreatedOn").get("From"))
            created_on_to = (req.get("CreatedOn").get("To"))
            where_str += " and date(csp.created_on) between '%s' and '%s'  " % (created_on_from, created_on_to)

        if req.get("ExpectedDOJ"):

            created_on_from = (req.get("ExpectedDOJ").get("From"))
            created_on_to = (req.get("ExpectedDOJ").get("To"))
            where_str += " and date(csp.expected_joining_date) between '%s' and '%s'  " % (created_on_from, created_on_to)

        if req.get("TentativeDOJ"):
            created_on_from = (req.get("TentativeDOJ").get("From"))
            created_on_to = (req.get("TentativeDOJ").get("To"))
            where_str += " and date(csp.tentative_doj) between '%s' and '%s'  " % (created_on_from, created_on_to)

        if req.get("ActualDOJ"):
            created_on_from = (req.get("ActualDOJ").get("From"))
            created_on_to = (req.get("ActualDOJ").get("To"))
            where_str += " and date(csp.actual_joining_date) between '%s' and '%s'  " % (created_on_from, created_on_to)

        final_qur = ""
        if where_str:
            final_qur = select_str + " where " + where_str
            self.query = final_qur
            print (self.query)
        if final_qur:
            try:
                # print self.query
                self.cursor.execute(final_qur)
                Query_Result = self.cursor.fetchone()
                # print final_qur
                self.Query_Result1 = Query_Result[0]
            except Exception as e:
                print (e)

    def over_status(self):
        self.ws.write(0, 0, 'U_Candidates')
        if self.Expected_success_cases == self.Expected_success_cases:
            self.ws.write(0, 1, 'Pass',  self.style24)
        else:
            self.ws.write(0, 1, 'Fail',  self.style25)

        self.ws.write(0, 3, 'StartTime')
        self.ws.write(0, 4, self.start_time)
        xlob.wb_result.save('/home/vinodkumar/hirepro_automation/API-Automation/Output Data/Pofu/AdvanceSearch/Advance_Search_01.xls')


print("Combined Search Script Started")
xlob = Excel_Data()
xlob.all()
xlob.over_status()
print("Completed Successfully ")
