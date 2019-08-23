from hpro_automation import (login, db_login, input_paths, output_paths, work_book)
import dateutil.parser
import requests
import json
import copy
import xlrd
import xlwt
import datetime


class ExcelData(login.CommonLogin, work_book.WorkBook, db_login.DBConnection):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(ExcelData, self).__init__()
        self.common_login('crpo')
        self.db_connection('amsin')

        # This Script works for below fields
        self.xl_json_request = []
        self.xl_expected_candidate_id = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 100)))
        self.Actual_Success_case = []
        self.success_case_01 = {}
        self.success_case_02 = {}

        self.rownum = 2

        self.event_filter = ["CandidateName","CandidateIds","ApplicantIds","Email","Usn","Phone","GenderType",
                            "CreatedFrom","CreatedTo","ModifiedFrom","ModifiedTo","CreatedByIds","ModifiedByIds",
                            "jobRoleId","BusinessUnitIds","ResumeStatusIds","CommunicationPurpose","CommunicationStatus",
                            "CollegeIds","DegreeIds","BranchIds","YearOfPassing","PercentageFrom","PercentageTo","IsFinal",
                            "TestId","OnlineAssessmentStatus","OnlineAssessmentMarksFrom","OnlineAssessmentMarksTo",
                            "TestAttendedFrom","TestAttendedTo","TaskIds","TaskStatusType","ActivityIds","ActivityStatusType",
                            "Text1","Text2","Text3","Text4","Text5","Text6","Text7","Text8","Text9","Text10","Text11",
                            "Text12","Text13","Text14","Text15","Integer1Ids","Integer2Ids","Integer3Ids","Integer4Ids",
                            "Integer5Ids","Integer6Ids","Integer7Ids","Integer8Ids","Integer9Ids","Integer10Ids",
                            "Integer11Ids","Integer12Ids","Integer13Ids","Integer14Ids","Integer15Ids","TextArea1",
                            "TextArea2","TextArea3","TextArea4","TrueFalse1","TrueFalse2","TrueFalse3","TrueFalse4","TrueFalse5",
                            "DateCustomField1From","DateCustomField1To","DateCustomFie2d1From","DateCustomField2To",
                            "DateCustomField3From","DateCustomField3To","DateCustomField4From","DateCustomField4To",
                            "DateCustomField5From","DateCustomField5To","IsApplicantHistory"]

        self.job_filter = ["JobId"]
        self.requirement_filter = ["RecruitEventId"]
        self.our_expected_ids = [1219152,1219153,1219154,1219155,1219156,1219157,1219158,1219159,1219160,1219161,
                                 1219162,1219163,1219164,1219165,1219166,1219167,1219168,1219169,1219170,1219171,
                                 1219172,1219173,1219174,1219175,1219176,1219177,1219178,1219179,1219180,1219181,
                                 1219182,1219183,1219184,1219185,1219186,1219187,1219188,1219189,1219190,1219191,
                                 1219192,1219193,1219194,1219195,1219196,1219197,1219198,1219199,1219200,1219201]
        # --------------------------------------------------------------------------------------------------------------
        # Write Excel Header's
        # --------------------------------------------------------------------------------------------------------------
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Search Result')
        self.ws.write(1, 0, 'Request', self.style0)
        self.ws.write(1, 1, 'Request', self.style0)
        self.ws.write(1, 2, 'API Count', self.style0)
        self.ws.write(1, 3, 'DB Count', self.style0)
        self.ws.write(1, 4, 'Expected Candidate Id\'s', self.style0)
        self.ws.write(1, 5, 'Not Matched Id\'s', self.style0)

        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")

    # ------------------------------------------------------------------------------------------------------------------
    # Read Excel Data
    # ------------------------------------------------------------------------------------------------------------------
        wb = xlrd.open_workbook(input_paths.inputpaths['Applicant_count_Input_sheet'])
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        # print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range(1, sh1.nrows):
            rownum = i
            rows = sh1.row_values(rownum)
            self.xl_json_request.append(rows[0])
            self.xl_expected_candidate_id.append(str(rows[1]))

        local = self.xl_expected_candidate_id
        length = len(self.xl_expected_candidate_id)
        self.new_local = []
        # --------------------------------------------------------------------------------------------------------------
        # below loop is used to convert input expecetd id's in to lists
        # Note:- for comparisition we need lists
        # --------------------------------------------------------------------------------------------------------------
        for i in range(0, length):
            j = [int(float(b)) for b in local[i].split(',')]
            self.new_local.append(j)
        self.xl_expected = self.new_local

    # ------------------------------------------------------------------------------------------------------------------
    # below method is used to get count from the api response
    # ------------------------------------------------------------------------------------------------------------------
    def json_data(self):

        self.lambda_function('getAllApplicants')
        self.headers['APP-NAME'] = 'crpo'

        r = requests.post(self.webapi, headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
        print(r.headers)
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']

        if self.status == 'OK':
            self.count = resp_dict['TotalItemCount']
            self.total_pages = int(resp_dict['TotalPages'])
            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count

    # ------------------------------------------------------------------------------------------------------------------
    # below method is used to get all candidate id's from the api response
    # ------------------------------------------------------------------------------------------------------------------
    def json_data_iteration(self, data, iter):

        self.lambda_function('getAllApplicants')
        self.headers['APP-NAME'] = 'crpo'

        iter += 1
        self.actual_ids = []
        for i in range(1, iter):
            self.data["PagingCriteria"]["PageNo"] = i
            r = requests.post(self.webapi, headers=self.headers, data=json.dumps(data, default=str), verify=False)
            print(r.headers)
            resp_dict = json.loads(r.content)
            # print resp_dict
            for element in resp_dict["data"]:
                self.actual_ids.append(element["CandidateId"])
                # print element1
        # print len(self.actual_ids)
        # print self.actual_ids

    def all(self):
        tot_len = len(self.xl_json_request)
        for i in range(0, tot_len):
            print("Iteration Count :- %s " %i)

            # ----------------------------------------------------------------------------------------------------------
            # Usually Data comes as unicode format from Excel so we need to convert in to dictionary
            # Here json.loads is used to convert Unicode to Dictionary
            # ----------------------------------------------------------------------------------------------------------
            self.xl_request=json.loads(self.xl_json_request[i])
            # print self.xl_request

            # this sis Our Original input Request
            self.xl_request1 = copy.deepcopy(self.xl_request)

            # print self.xl_request1

            if self.xl_request.get("CandidateIds"):
                self.xl_request["CandidateIds"] = self.our_expected_ids
                # print self.xl_request
            else:
                val = [("CandidateIds",self.our_expected_ids)]
                id_filter = dict(val)
                self.xl_request.update(id_filter)
                # print self.xl_request

            # ----------------------------------------------------------------------------------------------------------
            # keys() method is used to get all the keyword's from the request
            # ----------------------------------------------------------------------------------------------------------
            all_keys = self.xl_request .keys()
            # print all_keys

            # ----------------------------------------------------------------------------------------------------------
            # below lines are used to segregate all the keywords
            # ----------------------------------------------------------------------------------------------------------
            event_requests = set(all_keys) & set(self.event_filter)
            event_requests = list(event_requests)
            # print candidate_requests

            job_requests = set(all_keys) & set(self.job_filter)
            job_requests = list(job_requests)
            # print applicant_requests

            requirement_requests = set(all_keys) & set(self.requirement_filter)
            requirement_requests = list(requirement_requests)

            # ----------------------------------------------------------------------------------------------------------
            # Making Json fromat
            # ----------------------------------------------------------------------------------------------------------
            final_event_filter_request = []
            total = len(event_requests)
            for j in range(0, total):
                all_keys = (event_requests[j], self.xl_request[event_requests[j]])
                # print all_keys
                final_event_filter_request.append(all_keys)
            final_event_filter_request = dict(final_event_filter_request)
            # print  event_filter_request

            total = len(job_requests)
            if total==0  or total == None :
                final_job_filter_request = ""
            else:
                final_job_filter_request = self.xl_request["JobId"]

            total = len(requirement_requests)
            if total == 0 or total == None :
                final_requirement_filter_request = ""
            else:
                final_requirement_filter_request = self.xl_request["RecruitEventId"]


            # print custom_requests
            # ----------------------------------------------------------------------------------------------------------
            # Search API Call
            # ----------------------------------------------------------------------------------------------------------
            # self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.get_token}
            self.data = {
                "PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 200, "PageNo": 1,
                                   "SortParameter": "0", "SortOrder": "0", "PropertyIds": [], "ObjectState": 0,
                                   "IsCountOnly": True}, "SearchCampusApplicantType":final_event_filter_request,
                                    "RecruitEventId":final_requirement_filter_request,"JobId":final_job_filter_request,
                                    "IsNotCacheRequired": False}
            # print self.data
            self.json_data()
            self.total_api_count = self.count
            if self.count != "400000000000000":
                self.data["PagingCriteria"] = {"IsRefresh": False, "MaxResults": 200, "PageNo": 1, "ObjectState": 0}
                # print self.data
                self.json_data_iteration(self.data, self.total_pages)
            self.mismatched_id = set(self.xl_expected[i]) - set(self.actual_ids)

            self.Query_Generation()

            expected_id = str(self.xl_expected[i])
            expected_id = expected_id.strip('[]')
            mismatched_id = str(list(self.mismatched_id))
            mismatched_id = mismatched_id.strip('[]')

            # ----------------------------------------------------------------------------------------------------------
            # Write Excel Method
            # ----------------------------------------------------------------------------------------------------------
            # print self.xl_request1
            if self.total_api_count == self.Query_Result1:
                self.ws.write(self.rownum, 0, 'Pass', self.style26)
                self.success_case_01 = 'Pass'
            elif self.Query_Result1 == 2:
                self.ws.write(self.rownum, 0, 'Pass', self.style26)
                self.success_case_02 = 'Pass'
            else:
                self.ws.write(self.rownum, 0, 'Fail', self.style3)

            if self.success_case_01 == 'Pass':
                self.Actual_Success_case.append(self.success_case_01)
            if self.success_case_02 == 'Pass':
                self.Actual_Success_case.append(self.success_case_02)

            Object.success_case_01 = {}
            Object.success_case_02 = {}

            self.ws.write(self.rownum, 1, str(self.xl_request1))
            if self.total_api_count == self.Query_Result1:
                self.ws.write(self.rownum, 2, self.total_api_count, self.style12)
                self.ws.write(self.rownum, 3, self.Query_Result1, self.style14)
                self.ws.write(self.rownum, 4, expected_id, self.style14)
                self.ws.write(self.rownum, 5, mismatched_id, self.style3)

            elif self.total_api_count == '400000000000000':
                print("API Failed")
                self.ws.write(self.rownum, 2, "API Failed", self.style3)
                self.ws.write(self.rownum, 3, self.Query_Result1, self.style12)
                self.ws.write(self.rownum, 4, expected_id, self.style3)
                self.ws.write(self.rownum, 5, "API Failed", self.style3)
            else:
                print("this is else part \ n")
                self.ws.write(self.rownum, 2, self.total_api_count, self.style12)

                if self.Query_Result1 == 2:
                    self.ws.write(self.rownum, 3, self.Query_Result1, self.style7)
                else:
                    self.ws.write(self.rownum, 3, self.Query_Result1, self.style3)

                if self.Query_Result1 == 2:
                    self.ws.write(self.rownum, 4, expected_id, self.style7)
                else:
                    self.ws.write(self.rownum, 4, expected_id, self.style3)

                self.ws.write(self.rownum, 5, mismatched_id, self.style3)
                print(self.query)
            self.wb_result.save(output_paths.outputpaths['Applicant_count_Output_sheet_2'])
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1

    def Query_Generation(self):
        select_str = "select count(distinct(c.id)) from candidates c " \
                     "left join candidate_education_profiles ce on c.id = ce.candidate_id " \
                     "left join candidate_customs cc on c.candidatecustom_id = cc.id "\
                     "left join applicant_statuss sp on c.id = sp.candidate_id "\
                     "left join test_users tu on c.id=tu.candidate_id "\
                     "left join master_job_requisitions mjr on mjr.id=sp.master_job_requisition_id "\
                    "left join assigned_tasks at on at.candidate_id = c.id "\
                    "left join applicant_status_items asi on sp.id=asi.applicantstatus_id "
        where_str = " c.is_deleted=0 and c.is_draft=0 and c.is_archived=0 and c.tenant_id=1787 and sp.is_deleted=0 "

        if self.xl_request.get("CandidateName"):
            # print self.xl_request.get("EventId")
            where_str += " and c.candidate_name like '%{}%' ".format(self.xl_request.get("CandidateName"))

        if self.xl_request.get("ApplicantIds"):
            a = self.xl_request.get("ApplicantIds")
            values = ','.join(str(v) for v in a)
            where_str += " and sp.id in ( %s ) "%values

        if self.xl_request.get("CandidateIds"):
            a = self.xl_request.get("CandidateIds")
            values = ','.join(str(v) for v in a)
            event_id =  self.xl_request.get("eventId")
            where_str += " and sp.candidate_id in ( %s ) and sp.recruitevent_id= %s " %(values,event_id)

        if self.xl_request.get("Email"):
            where_str += " and c.email1 like '%{}%' ".format(self.xl_request.get("Email"))

        if self.xl_request.get("Usn"):
            where_str += " and c.usn = '{}' ".format(self.xl_request.get("Usn"))

        if self.xl_request.get("Phone"):
            where_str += " and c.mobile1 = {} ".format(self.xl_request.get("Phone"))

        if self.xl_request.get("GenderType"):
            where_str += " and c.gender = {} ".format(self.xl_request.get("GenderType"))

        if self.xl_request.get("CreatedByIds"):
            where_str += " and sp.created_by in ( {} ) ".format(self.xl_request.get("CreatedByIds")[0])

        if self.xl_request.get("ModifiedByIds"):
            where_str += " and sp.modified_by in ( {} ) ".format(self.xl_request.get("ModifiedByIds")[0])

        if self.xl_request.get("CreatedFrom") and self.xl_request.get("CreatedTo"):
            tagged_on_from = self.xl_request.get("CreatedFrom")
            tagged_on_to = self.xl_request.get("CreatedTo")
            where_str += " and sp.created_on between '%s' and '%s' " %(tagged_on_from,tagged_on_to)

        if self.xl_request.get("ModifiedFrom") and self.xl_request.get("ModifiedTo"):
            modified_on_from = self.xl_request.get("ModifiedFrom")
            modified_on_to = self.xl_request.get("ModifiedTo")
            where_str += " and sp.modified_on between '%s' and '%s' "% (modified_on_from, modified_on_to)

        if self.xl_request.get("jobRoleId"):
            where_str += " and sp.job_id = {} ".format(self.xl_request.get("jobRoleId"))

        if self.xl_request.get("IsFinal") == 0:

            if self.xl_request.get("CollegeIds"):
                where_str += " and  ce.college_id in ({}) ".format(self.xl_request.get("CollegeIds")[0])

            if self.xl_request.get("DegreeIds"):
                where_str += " and  ce.degree_id in ({}) ".format(self.xl_request.get("DegreeIds")[0])

            if self.xl_request.get("BranchIds"):
                where_str += " and  ce.degree_type_id  in({}) ".format(self.xl_request.get("BranchIds")[0])

            if self.xl_request.get("YearOfPassing"):
                where_str += " and  ce.end_year in ({}) ".format(self.xl_request.get("YearOfPassing"))

            if self.xl_request.get("PercentageTo"):
                percentage_from = self.xl_request.get("PercentageFrom")
                percentage_to = self.xl_request.get("PercentageTo")
                where_str += " and ce.percentage  between '%s' and '%s' "%(percentage_from,percentage_to)
        else:
            if self.xl_request.get("CollegeIds"):
                where_str += " and  ce.college_id in ({}) and ce.is_final=1".format(self.xl_request.get("CollegeIds")[0])

            if self.xl_request.get("DegreeIds"):
                where_str += " and  ce.degree_id in ({}) and ce.is_final=1".format(self.xl_request.get("DegreeIds")[0])

            if self.xl_request.get("BranchIds"):
                where_str += " and  ce.degree_type_id  in({}) and ce.is_final=1".format(self.xl_request.get("BranchIds")[0])

            if self.xl_request.get("YearOfPassing"):
                where_str += " and  ce.end_year in ({}) and ce.is_final=1".format(self.xl_request.get("YearOfPassing"))

            if self.xl_request.get("PercentageTo"):
                percentage_from = self.xl_request.get("PercentageFrom")
                percentage_to = self.xl_request.get("PercentageTo")
                where_str += " and ce.percentage  between '%s' and '%s' and ce.is_final=1"% (percentage_from, percentage_to)

        if self.xl_request.get("Integer1Ids"):
            where_str += " and  c.integer1 =  {} ".format(self.xl_request.get("Integer1Ids")[0])

        if self.xl_request.get("Integer2Ids"):
            where_str += " and  c.integer2 =  {} ".format(self.xl_request.get("Integer2Ids")[0])

        if self.xl_request.get("Integer3Ids"):
            where_str += " and  c.integer3 =  {} ".format(self.xl_request.get("Integer3Ids")[0])

        if self.xl_request.get("Integer4Ids"):
            where_str += " and  c.integer4 =  {} ".format(self.xl_request.get("Integer4Ids")[0])

        if self.xl_request.get("Integer5Ids"):
            where_str += " and  c.integer5 =  {} ".format(self.xl_request.get("Integer5Ids")[0])

        if self.xl_request.get("Integer6Ids"):
            where_str += " and  cc.integer6 =  {} ".format(self.xl_request.get("Integer6Ids")[0])

        if self.xl_request.get("Integer7Ids"):
            where_str += " and  cc.integer7 =  {} ".format(self.xl_request.get("Integer7Ids")[0])

        if self.xl_request.get("Integer8Ids"):
            where_str += " and  cc.integer8 =  {} ".format(self.xl_request.get("Integer8Ids")[0])

        if self.xl_request.get("Integer9Ids"):
            where_str += " and  cc.integer9 =  {} ".format(self.xl_request.get("Integer9Ids")[0])

        if self.xl_request.get("Integer10Ids"):
            where_str += " and  cc.integer10 =  {} ".format(self.xl_request.get("Integer10Ids")[0])

        if self.xl_request.get("Integer11Ids"):
            where_str += " and  cc.integer11 =  {} ".format(self.xl_request.get("Integer11Ids")[0])

        if self.xl_request.get("Integer12Ids"):
            where_str += " and  cc.integer12 =  {} ".format(self.xl_request.get("Integer12Ids")[0])

        if self.xl_request.get("Integer13Ids"):
            where_str += " and  cc.integer13 =  {} ".format(self.xl_request.get("Integer13Ids")[0])

        if self.xl_request.get("Integer14Ids"):
            where_str += " and  cc.integer14 =  {} ".format(self.xl_request.get("Integer14Ids")[0])

        if self.xl_request.get("Integer15Ids"):
            where_str += " and  cc.integer15 =  {} ".format(self.xl_request.get("Integer15Ids")[0])


        if self.xl_request.get("Text1"):
            where_str += " and  c.text1 like  '%{}%' ".format(self.xl_request.get("Text1"))

        if self.xl_request.get("Text2"):
            where_str += " and  c.text2 like  '%{}%' ".format(self.xl_request.get("Text2"))

        if self.xl_request.get("Text3"):
            where_str += " and  c.text3 like  '%{}%' ".format(self.xl_request.get("Text3"))

        if self.xl_request.get("Text4"):
            where_str += " and  c.text4 like '%{}%' ".format(self.xl_request.get("Text4"))

        if self.xl_request.get("Text5"):
            where_str += " and  c.text5 like '%{}%' ".format(self.xl_request.get("Text5"))

        if self.xl_request.get("Text6"):
            where_str += " and  cc.text6 like '%{}%' ".format(self.xl_request.get("Text6"))

        if self.xl_request.get("Text7"):
            where_str += " and  cc.text7 like '%{}%' ".format(self.xl_request.get("Text7"))

        if self.xl_request.get("Text8"):
            where_str += " and  cc.text8 like '%{}%' ".format(self.xl_request.get("Text8"))

        if self.xl_request.get("Text9"):
            where_str += " and  cc.text9 like '%{}%' ".format(self.xl_request.get("Text9"))

        if self.xl_request.get("Text10"):
            where_str += " and  cc.text10 like '%{}%' ".format(self.xl_request.get("Text10"))

        if self.xl_request.get("Text11"):
            where_str += " and  cc.text11 like '%{}%' ".format(self.xl_request.get("Text11"))

        if self.xl_request.get("Text12"):
            where_str += " and  cc.text12 like '%{}%' ".format(self.xl_request.get("Text12"))

        if self.xl_request.get("Text13"):
            where_str += " and  cc.text13 like '%{}%' ".format(self.xl_request.get("Text13"))

        if self.xl_request.get("Text14"):
            where_str += " and  cc.text14 like '%{}%' ".format(self.xl_request.get("Text14"))

        if self.xl_request.get("Text15"):
            where_str += " and  cc.text15 like '%{}%' ".format(self.xl_request.get("Text15"))

        if self.xl_request.get("TextArea1"):
            where_str += " and  c.text_area1 like '%{}%' ".format(self.xl_request.get("TextArea1"))

        if self.xl_request.get("TextArea2"):
            where_str += " and  c.text_area2 like '%{}%' ".format(self.xl_request.get("TextArea2"))

        if self.xl_request.get("TextArea3"):
            where_str += " and  c.text_area3 like '%{}%' ".format(self.xl_request.get("TextArea3"))

        if self.xl_request.get("TextArea4"):
            where_str += " and  c.text_area4 like '%{}%' ".format(self.xl_request.get("TextArea4"))

        if self.xl_request.get("DateCustomField1From"):
            date1 =self.xl_request.get("DateCustomField1From")
            # event_id = self.xl_request.get("eventId")
            # print date1
            where_str += " and date(c.date_custom_field1) = '%s' " %(date1)

        if self.xl_request.get("DateCustomField2From"):
            date2 = self.xl_request.get("DateCustomField2From")
            # print date2
            where_str += " and date(c.date_custom_field2) = '%s' " %(date2)

        if self.xl_request.get("DateCustomField3From"):
            date3 = self.xl_request.get("DateCustomField3From")
            # print date3
            where_str += " and date(cc.date_custom_field3) = '%s' " %(date3)


        if self.xl_request.get("DateCustomField4From"):
            date4 = self.xl_request.get("DateCustomField4From")
            # print date4
            where_str += " and date(cc.date_custom_field4) = '%s' " %(date4)


        if self.xl_request.get("DateCustomField5From"):
            date5 = self.xl_request.get("DateCustomField5From")
            where_str += " and date(cc.date_custom_field5) = '%s' " %(date5)


        if self.xl_request.get("TrueFalse1") in [0,1]:
            date1 = self.xl_request.get("TrueFalse1")
            # print date1
            where_str += " and  c.true_false1 = %s " %(date1)

        if self.xl_request.get("TrueFalse2") in [0,1]:
            date2 =self.xl_request.get("TrueFalse2")
            # print date2
            where_str += " and  c.true_false2 = %s " %(date2)

        if self.xl_request.get("TrueFalse3") in [0,1]:
            date3 = self.xl_request.get("TrueFalse3")
            # print date3
            where_str += " and  cc.true_false3 = %s " %(date3)


        if self.xl_request.get("TrueFalse4") in [0,1]:
            date4 = self.xl_request.get("TrueFalse4")
            # print date4
            where_str += " and  cc.true_false4 = %s " %(date4)


        if self.xl_request.get("TrueFalse5") in [0,1]:
            date5 = self.xl_request.get("TrueFalse5")
            where_str += " and  cc.true_false5 = %s " %(date5)

        if self.xl_request.get("TestId"):
            where_str += " and tu.test_id =  {}".format(self.xl_request.get("TestId"))

        if self.xl_request.get("OnlineAssessmentStatus")==0 or self.xl_request.get("OnlineAssessmentStatus")==1 :
            where_str += " and tu.test_id in (select test_id from recruit_event_job_tests where recruitevent_id = {} )" \
                         " and tu.status =  {}".format(self.xl_request.get("eventId"),
                                                       self.xl_request.get("OnlineAssessmentStatus"))

        # if self.xl_request.get("credential"):
        #     where_str += " and tu.test_id in (select test_id from recruit_event_job_tests where recruitevent_id= {})" \
        #                  " and tu.status =  {}".format(self.xl_request.get("eventId"),
        #                                                self.xl_request.get("OnlineAssessmentStatus")[0])

        if self.xl_request.get("OnlineAssessmentMarksFrom") and self.xl_request.get("OnlineAssessmentMarksTo"):
            testscore_from = self.xl_request.get("OnlineAssessmentMarksFrom")
            testscore_to = self.xl_request.get("OnlineAssessmentMarksTo")
            event_id = self.xl_request.get("eventId")
            where_str += " and tu.test_id in (select test_id from recruit_event_job_tests where recruitevent_id= %s)" \
                         " and tu.total_score  between '%s' and '%s' " %(event_id,testscore_from,testscore_to)

        if self.xl_request.get("TestAttendedFrom") and self.xl_request.get("TestAttendedTo"):
            attended_on_from = self.xl_request.get("TestAttendedFrom")
            attended_on_to = self.xl_request.get("TestAttendedTo")
            event_id =  self.xl_request.get("eventId")
            where_str += " and tu.login_time is not null and tu.test_id in (select test_id from recruit_event_job_tests " \
                         "where recruitevent_id= %s) and login_time between '%s' and '%s' " \
                         " " % (event_id, attended_on_from, attended_on_to)

        if self.xl_request.get("BusinessUnitIds"):
            where_str += " and mjr.job_id={} and mjr.business_unit_id={} "\
                .format(self.xl_request.get("jobRoleId"),self.xl_request.get("BusinessUnitIds")[0])

        if self.xl_request.get("IsApplicantHistory")==1:
            if self.xl_request.get("ResumeStatusIds"):
                where_str += " and asi.status_id={} and sp.job_id={} " \
                    .format(self.xl_request.get("ResumeStatusIds")[0], self.xl_request.get("jobRoleId"))
        else:

            if self.xl_request.get("ResumeStatusIds"):
                where_str += " and sp.current_status_id = {} and sp.job_id={} "\
                    .format(self.xl_request.get("ResumeStatusIds")[0],self.xl_request.get("jobRoleId"))

        if self.xl_request.get("ActivityIds"):
            where_str += " and c.current_activity ={} " \
                .format(self.xl_request.get("ActivityIds")[0])

        if self.xl_request.get("ActivityStatusType"):
            where_str += " and c.current_activity ={} and c.activity_status={}" \
                .format(self.xl_request.get("ActivityIds")[0],self.xl_request.get("ActivityStatusType"))

        if self.xl_request.get("TaskIds"):
            where_str += " and at.candidate_id in(select candidate_id from applicant_statuss where recruitevent_id={})" \
                         " and at.task_id={} and at.status ={} " \
                .format(self.xl_request.get("RecruitEventId"),self.xl_request.get("TaskIds")[0],self.xl_request.get("TaskStatusType"))

        # candidate_filter = " and c.is_deleted=0 and c.is_draft=0 and c.is_archived=0 and c.tenant_id=1787;"
        final_qur = ""
        if where_str:
            final_qur = select_str + " where " + where_str
            self.query = final_qur
        if final_qur:
            try:
                # print self.query
                self.cursor.execute(final_qur)
                Query_Result = self.cursor.fetchone()
                # print final_qur
                self.Query_Result1 = Query_Result[0]
            except Exception as e:
                print(e)

    def date_converter(self, input_date):
        converted_utc_date = dateutil.parser.parse(input_date)
        converted_local_date = converted_utc_date.astimezone(dateutil.tz.tzlocal()).replace(tzinfo=None)
        date = converted_local_date.strftime("%Y-%m-%d")
        return date

    def overall_status(self):
        self.ws.write(0, 0, 'Applicant Boundary Search', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, 'in progress...', self.style24)
        Object.wb_result.save(output_paths.outputpaths['Applicant_count_Output_sheet_2'])


print("Combined Search Script Started")
Object = ExcelData()
Object.all()
print("Completed Successfully ")
Object.overall_status()
