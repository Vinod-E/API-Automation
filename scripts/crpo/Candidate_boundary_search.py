import copy
import dateutil.parser
import requests
import json
import xlrd
import xlwt
import datetime
from hpro_automation import (login, api, output_paths, input_paths, db_login, work_book)


class ExcelData(login.CommonLogin, db_login.DBConnection, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(ExcelData, self).__init__()
        self.common_login('crpo')
        self.db_connection('amsin')

        # This Script works for below fields
        self.candidate_filter = ['CandidateIds', 'Name', 'Email', 'Phone', 'USN', 'CreatedBy', 'PanNo', 'AadhaarNo',
                                 'CurrentLocationIds', 'TypesOfSource', 'SourceId', 'Sourcer', 'Gender',
                                 'MaritalStatus',
                                 'Gender', 'MaritalStatus', 'CollegeIds', 'DegreeIds', 'BranchIds',
                                 'YearEnd', 'ExpertiseIds', 'WorkProfileOrganisationIds', 'Text1', 'Text2', 'Text3',
                                 'Text4', 'Text5', 'Integer1', 'Integer2', 'Integer3', 'Integer4', 'Integer5',
                                 'TextArea1', 'TextArea2', 'TextArea3', 'TextArea4', 'TextArea5', 'DateCustomField1',
                                 'DateCustomField2', 'CandidateUtilization', 'ExperienceFrom', 'ExperienceTo',
                                 'PercentageFrom', 'PercentageTo', 'CreatedFrom', 'CreatedTo', 'ModifiedFrom',
                                 'ModifiedTo']

        self.applicant_filters = ['JobDepartmentIds', 'JobIds', 'ApplicantStatusIds', 'EventIds']

        self.candidate_custom_filters = ['Text6', 'Text7', 'Text8', 'Text9', 'Text10', 'Text11', 'Text12', 'Text13',
                                         'Text14', 'Text15', 'Integer6', 'Integer7', 'Integer8', 'Integer9',
                                         'Integer10', 'Integer11', 'Integer12', 'Integer13', 'Integer14', 'Integer15',
                                         'DateCustomField3', 'DateCustomField4',  'DateCustomField5']
        self.candidate_test_filter = ["TestIds", "StatusIds", "TestScoreFrom", "TestScoreTo"]
        self.our_expected_ids = [1219152, 1219153, 1219154, 1219155, 1219156, 1219157, 1219158, 1219159, 1219160,
                                 1219161, 1219162, 1219163, 1219164, 1219165, 1219166, 1219167, 1219168, 1219169,
                                 1219170, 1219171, 1219172, 1219173, 1219174, 1219175, 1219176, 1219177, 1219178,
                                 1219179, 1219180, 1219181, 1219182, 1219183, 1219184, 1219185, 1219186, 1219187,
                                 1219188, 1219189, 1219190, 1219191, 1219192, 1219193, 1219194, 1219195, 1219196,
                                 1219197, 1219198, 1219199, 1219200, 1219201]

        self.xl_json_request = []
        self.xl_expected_candidate_id = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 96)))
        self.Actual_Success_case = []
        self.success_case_01 = {}
        self.rownum = 2
        # --------------------------------------------------------------------------------------------------------------
        # Write Excel Header's
        # --------------------------------------------------------------------------------------------------------------
        self.wb_result = xlwt.Workbook()
        self.ws = self.wb_result.add_sheet('Candidate Search Result')
        self.ws.write(1, 0, 'Actual_Status', self.style0)
        self.ws.write(1, 1, 'Request', self.style0)
        self.ws.write(1, 2, 'API Count', self.style0)
        self.ws.write(1, 3, 'DB Count', self.style0)
        self.ws.write(1, 4, 'Expected Candidate Id\'s', self.style0)
        self.ws.write(1, 5, 'Not Matched Id\'s', self.style0)
        # self.ws.write(0, 5, 'API Id\'s', self.style11)

        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")

    # ------------------------------------------------------------------------------------------------------------------
    # Read Excel Data
    # ------------------------------------------------------------------------------------------------------------------
        wb = xlrd.open_workbook(input_paths.inputpaths['candidate_search_Input_sheet'])
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

    def json_data(self):

        self.lambda_function('get_all_candidates')
        self.headers['APP-NAME'] = 'crpo'

        r = requests.post(self.webapi, headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
        print(r.headers)
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        print(resp_dict)

        if self.status == 'OK':
            self.count = resp_dict['TotalItem']
            if resp_dict['TotalPages']!=None:
                self.total_pages = int(resp_dict['TotalPages'])
            else:
                self.total_pages = 1
            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count

    def json_data_iteration(self, data, iter):

        self.lambda_function('get_all_candidates')
        self.headers['APP-NAME'] = 'crpo'

        iter += 1
        self.actual_ids = []
        for i in range(1, iter):
            self.data["PagingCriteria"]["PageNo"] = i
            r = requests.post(self.webapi, headers=self.headers, data=json.dumps(data, default=str), verify=False)
            print(r.headers)
            resp_dict = json.loads(r.content)
            # print resp_dict
            for element in resp_dict["Candidates"]:
                self.actual_ids.append(element["Id"])
                # print element1
        # print len(self.actual_ids)
        # print self.actual_ids

    def all(self):
        self.tot_len = len(self.xl_json_request)
        for i in range(0, self.tot_len):
            print("Iteration Count :- %s " %i)
            final_candidate_filter_request = []
            final_applicant_filter_request = []
            final_custom_filter_request = []
            final_candidate_test_filter_request = []
            # ----------------------------------------------------------------------------------------------------------
            # Usually Data comes as unicode format from Excel so we need to convert in to dictionary
            # Here ast.literal_eval is used to convert Unicode to Dictionary
            # ----------------------------------------------------------------------------------------------------------
            # self.xl_request = ast.literal_eval(self.xl_json_request[i])
            self.xl_request = json.loads(self.xl_json_request[i])
            self.xl_request1 = copy.deepcopy(self.xl_request)

            if self.xl_request.get("CandidateIds"):
                self.xl_request["CandidateIds"] = self.our_expected_ids
                # print self.xl_request
            else:
                val = [("CandidateIds", self.our_expected_ids)]
                id_filter = dict(val)
                self.xl_request.update(id_filter)
                # print self.xl_request

            all_keys = self.xl_request.keys()
            candidate_requests = set(all_keys) & set(self.candidate_filter)
            candidate_requests = list(candidate_requests)
            # print candidate_requests

            applicant_requests = set(all_keys) & set(self.applicant_filters)
            applicant_requests = list(applicant_requests)
            # print applicant_requests

            custom_requests = set(all_keys) & set(self.candidate_custom_filters)
            custom_requests = list(custom_requests)

            test_requests = set(all_keys) & set(self.candidate_test_filter)
            test_requests = list(test_requests)
            # print test_requests

            total = len(test_requests)
            for j in range(0, total):
                all_keys = (test_requests[j], self.xl_request[test_requests[j]])
                final_candidate_test_filter_request.append(all_keys)
            final_candidate_test_filter_request = dict(final_candidate_test_filter_request)
            # print  final_candidate_test_filter_request

            total = len(candidate_requests)
            for j in range(0, total):
                all_keys = (candidate_requests[j], self.xl_request[candidate_requests[j]])
                final_candidate_filter_request.append(all_keys)
            final_candidate_filter_request = dict(final_candidate_filter_request)
            # print  final_candidate_filter_request

            total = len(applicant_requests)
            for j in range(0, total):
                all_keys = (applicant_requests[j], self.xl_request[applicant_requests[j]])
                final_applicant_filter_request.append(all_keys)
            applicant_dictionary = dict(final_applicant_filter_request)
            # print applicant_dictionary

            total = len(custom_requests)
            for j in range(0, total):
                all_keys = (custom_requests[j], self.xl_request[custom_requests[j]])
                final_custom_filter_request.append(all_keys)
            custom_dictionary = dict(final_custom_filter_request)
            # print custom_dictionary

            # print custom_requests
            # ----------------------------------------------------------------------------------------------------------
            # Search API Call
            # ----------------------------------------------------------------------------------------------------------
            self.data = {
                "PagingCriteria": {"IsRefresh": False, "IsSpecificToUser": False, "MaxResults": 200, "PageNo": 1,
                                   "SortParameter": "0", "SortOrder": "0", "PropertyIds": [], "ObjectState": 0,
                                   "IsCountOnly": True}, "CandidateFilters": final_candidate_filter_request,
                                   "CandidateCustomFilters": custom_dictionary,
                                   "ApplicantFilters": applicant_dictionary,
                                   "CandidateTestFilters": final_candidate_test_filter_request,
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
            # ----------------------------------------------------------------------------------------------------------
            # Generating Dynamic query based on combinations and execute
            # ----------------------------------------------------------------------------------------------------------
            expected_id = str(self.xl_expected[i])
            expected_id = expected_id.strip('[]')
            mismatched_id = str(list(self.mismatched_id))
            mismatched_id = mismatched_id.strip('[]')

            if self.total_api_count == self.Query_Result1:
                self.ws.write(self.rownum, 0, 'Pass', self.style26)
                self.success_case_01 = 'Pass'
            else:
                self.ws.write(self.rownum, 0, 'Fail', self.style3)

            if self.success_case_01 == 'Pass':
                self.Actual_Success_case.append(self.success_case_01)

            self.ws.write(self.rownum, 1, str(self.xl_request1))
            if self.total_api_count == self.Query_Result1:
                self.ws.write(self.rownum, 2, self.total_api_count)
                self.ws.write(self.rownum, 3, self.Query_Result1, self.style14)
                self.ws.write(self.rownum, 4, expected_id, self.style14)
                self.ws.write(self.rownum, 5, mismatched_id, self.style3)
                # self.ws.write(self.rownum, 5, self.query, self.style14)

            elif self.total_api_count == '400000000000000':
                print("API Failed")
                self.ws.write(self.rownum, 2, "API Failed", self.style13)
                self.ws.write(self.rownum, 3, self.Query_Result1, self.style13)
                self.ws.write(self.rownum, 4, expected_id, self.style14)
                self.ws.write(self.rownum, 5, "API Failed", self.style13)
                # self.ws.write(self.rownum, 5, self.query, self.style13)
                print(self.query)
            else:
                print("this is else part \ n")
                self.ws.write(self.rownum, 2, self.total_api_count, self.style13)
                self.ws.write(self.rownum, 3, self.Query_Result1, self.style13)
                self.ws.write(self.rownum, 4, expected_id, self.style14)
                self.ws.write(self.rownum, 5, mismatched_id, self.style3)
                # self.ws.write(self.rownum, 5, self.query, self.style13)
                print(self.query)
            self.wb_result.save(output_paths.outputpaths['candidate_search_output_sheet_1'])
            # print statusCode, " -- ", b
            self.rownum = self.rownum + 1

    def Query_Generation(self):
        select_str = "select count(distinct(c.id)) from candidates c " \
                     "left join candidate_education_profiles ce on c.id = ce.candidate_id " \
                     "left join candidate_customs cc on c.candidatecustom_id = cc.id "\
                     "left join applicant_statuss sp on c.id = sp.candidate_id "\
                     "left join test_users tu on c.id=tu.candidate_id "

        where_str = "c.is_deleted=0 and c.is_draft=0 and c.is_archived=0 and c.tenant_id=1787"

        if self.xl_request.get("ExperienceFrom") and self.xl_request.get("ExperienceTo"):
            exp_from = self.xl_request.get("ExperienceFrom")
            exp_to = self.xl_request.get("ExperienceTo")
            where_str += " and c.total_experience between %s*12 and %s*12 " % (exp_from, exp_to)

        if self.xl_request.get("PercentageFrom") and self.xl_request.get("PercentageTo"):
            percentage_from = self.xl_request.get("PercentageFrom")
            percentage_to = self.xl_request.get("PercentageTo")
            where_str += " and ce.percentage  between '%s' and '%s' " % (percentage_from, percentage_to)

        if self.xl_request.get("CreatedFrom") and self.xl_request.get("CreatedTo"):
            created_on_from = Object.date_converter(self.xl_request.get("CreatedFrom"))
            created_on_to = Object.date_converter(self.xl_request.get("CreatedTo"))
            where_str += " and date(c.created_on) between '%s' and '%s' " % (created_on_from, created_on_to)

        if self.xl_request.get("ModifiedFrom") and self.xl_request.get("ModifiedTo"):
            modified_on_from = Object.date_converter(self.xl_request.get("ModifiedFrom"))
            modified_on_to = Object.date_converter(self.xl_request.get("ModifiedTo"))
            where_str += " and date(c.modified_on)  between '%s' and '%s' " % (modified_on_from, modified_on_to)

        if self.xl_request.get("Name"):
            where_str += " and c.candidate_name like '%{}%' ".format(self.xl_request.get("Name"))
        if self.xl_request.get("Email"):
            where_str += " and c.email1 like '%{}%' ".format(self.xl_request.get("Email"))
        if self.xl_request.get("Phone"):
            where_str += " and c.mobile1 = {} ".format(self.xl_request.get("Phone"))
        if self.xl_request.get("CandidateIds"):
            a = self.xl_request.get("CandidateIds")
            values = ','.join(str(v) for v in a)
            where_str += " and c.id in ( %s ) "%values

        if self.xl_request.get("USN"):
            where_str += " and c.USN = '{}'".format(self.xl_request.get("USN"))
        if self.xl_request.get("CreatedBy"):
            where_str += " and c.created_by = {}".format(self.xl_request.get("CreatedBy"))

        if self.xl_request.get("CandidateUtilization"):
            if self.xl_request.get("CandidateUtilization") == 'false':
                is_utilized = 0
            else:
                is_utilized = 1
            where_str += " and c.is_utilized = %s" %is_utilized

        if self.xl_request.get("PanNo"):
            where_str += " and c.pan_card = '{}'".format(self.xl_request.get("PanNo"))
        if self.xl_request.get("AadhaarNo"):
            where_str += " and c.aadhaar_no = {}".format(self.xl_request.get("AadhaarNo"))

        if self.xl_request.get("CurrentLocationIds"):
            where_str += " and c.current_location_id in ( {} ) ".format(','.join(self.xl_request.get("CurrentLocationIds")))

        if self.xl_request.get("TypesOfSource"):
            where_str += " and c.original_source_id in((select id from sources where types_of_source={}))"\
                .format(self.xl_request.get("TypesOfSource"))
        if self.xl_request.get("SourceId"):
            where_str += " and c.original_source_id = {}".format(self.xl_request.get("SourceId"))
        if self.xl_request.get("Sourcer"):
            where_str += " and c.sourcer = {}".format(self.xl_request.get("Sourcer"))
        if self.xl_request.get("Gender"):
            where_str += " and c.gender = {}".format(self.xl_request.get("Gender"))
        if self.xl_request.get("MaritalStatus"):
            where_str += " and c.marital_status = {}".format(self.xl_request.get("MaritalStatus"))

        if self.xl_request.get("ApplicantStatusIds"):
            where_str += " and sp.current_status_id = {}".format(self.xl_request.get("ApplicantStatusIds")[0])

        if self.xl_request.get("JobDepartmentIds"):
            where_str += " and sp.job_id in (select id from jobs where department_id={})"\
                .format(self.xl_request.get("JobDepartmentIds")[0])

        if self.xl_request.get("JobIds"):
            where_str += " and sp.job_id =  {}".format(self.xl_request.get("JobIds")[0])

        if self.xl_request.get("CollegeIds"):
            where_str += " and  ce.college_id =  {}".format(self.xl_request.get("CollegeIds")[0])
        if self.xl_request.get("DegreeIds"):
            where_str += " and  ce.degree_id =  {}".format(self.xl_request.get("DegreeIds")[0])
        if self.xl_request.get("BranchIds"):
            where_str += " and  ce.degree_type_id =  {}".format(self.xl_request.get("BranchIds")[0])
        if self.xl_request.get("YearEnd"):
            where_str += " and  ce.end_year =  {}".format(self.xl_request.get("YearEnd")[0])

        if self.xl_request.get("WorkProfileOrganisationIds"):
            where_str += " and  c.current_employer_id=  {}".format(self.xl_request.get("WorkProfileOrganisationIds")[0])
        if self.xl_request.get("ExpertiseIds"):
            where_str += " and  c.expertise_id1 =  {}".format(self.xl_request.get("ExpertiseIds")[0])

        if self.xl_request.get("Integer1"):
            where_str += " and  c.integer1 =  {}".format(self.xl_request.get("Integer1")[0])
        if self.xl_request.get("Integer1"):
            where_str += " and  c.integer1 =  {}".format(self.xl_request.get("Integer1")[0])
        if self.xl_request.get("Integer2"):
            where_str += " and  c.integer2 =  {}".format(self.xl_request.get("Integer2")[0])
        if self.xl_request.get("Integer3"):
            where_str += " and  c.integer3 =  {}".format(self.xl_request.get("Integer3")[0])
        if self.xl_request.get("Integer4"):
            where_str += " and  c.integer4 =  {}".format(self.xl_request.get("Integer4")[0])
        if self.xl_request.get("Integer5"):
            where_str += " and  c.integer5 =  {}".format(self.xl_request.get("Integer5")[0])
        if self.xl_request.get("Integer6"):
            where_str += " and  cc.integer6 =  {}".format(self.xl_request.get("Integer6")[0])
        if self.xl_request.get("Integer7"):
            where_str += " and  cc.integer7 =  {}".format(self.xl_request.get("Integer7")[0])
        if self.xl_request.get("Integer8"):
            where_str += " and  cc.integer8 =  {}".format(self.xl_request.get("Integer8")[0])
        if self.xl_request.get("Integer9"):
            where_str += " and  cc.integer9 =  {}".format(self.xl_request.get("Integer9")[0])
        if self.xl_request.get("Integer10"):
            where_str += " and  cc.integer10 =  {}".format(self.xl_request.get("Integer10")[0])
        if self.xl_request.get("Integer11"):
            where_str += " and  cc.integer11 =  {}".format(self.xl_request.get("Integer11")[0])
        if self.xl_request.get("Integer12"):
            where_str += " and  cc.integer12 =  {}".format(self.xl_request.get("Integer12")[0])
        if self.xl_request.get("Integer13"):
            where_str += " and  cc.integer13 =  {}".format(self.xl_request.get("Integer13")[0])
        if self.xl_request.get("Integer14"):
            where_str += " and  cc.integer14 =  {}".format(self.xl_request.get("Integer14")[0])
        if self.xl_request.get("Integer15"):
            where_str += " and  cc.integer15 =  {}".format(self.xl_request.get("Integer15")[0])

        if self.xl_request.get("Text1"):
            where_str += " and  c.text1 like  '%{}%'".format(self.xl_request.get("Text1"))
        if self.xl_request.get("Text2"):
            where_str += " and  c.text2 like  '%{}%'".format(self.xl_request.get("Text2"))
        if self.xl_request.get("Text3"):
            where_str += " and  c.text3 like  '%{}%'".format(self.xl_request.get("Text3"))
        if self.xl_request.get("Text4"):
            where_str += " and  c.text4 like  '%{}%'".format(self.xl_request.get("Text4"))
        if self.xl_request.get("Text5"):
            where_str += " and  c.text5 like  '%{}%'".format(self.xl_request.get("Text5"))
        if self.xl_request.get("Text6"):
            where_str += " and  cc.text6 like  '%{}%'".format(self.xl_request.get("Text6"))
        if self.xl_request.get("Text7"):
            where_str += " and  cc.text7 like  '%{}%'".format(self.xl_request.get("Text7"))
        if self.xl_request.get("Text8"):
            where_str += " and  cc.text8 like  '%{}%'".format(self.xl_request.get("Text8"))
        if self.xl_request.get("Text9"):
            where_str += " and  cc.text9 like  '%{}%'".format(self.xl_request.get("Text9"))
        if self.xl_request.get("Text10"):
            where_str += " and  cc.text10 like  '%{}%'".format(self.xl_request.get("Text10"))
        if self.xl_request.get("Text11"):
            where_str += " and  cc.text11 like  '%{}%'".format(self.xl_request.get("Text11"))
        if self.xl_request.get("Text12"):
            where_str += " and  cc.text12 like  '%{}%'".format(self.xl_request.get("Text12"))
        if self.xl_request.get("Text13"):
            where_str += " and  cc.text13 like  '%{}%'".format(self.xl_request.get("Text13"))
        if self.xl_request.get("Text14"):
            where_str += " and  cc.text14 like  '%{}%'".format(self.xl_request.get("Text14"))
        if self.xl_request.get("Text15"):
            where_str += " and  cc.text15 like  '%{}%'".format(self.xl_request.get("Text15"))

        if self.xl_request.get("TextArea1"):
            where_str += " and  c.text_area1 like '%{}%'".format(self.xl_request.get("TextArea1"))
        if self.xl_request.get("TextArea2"):
            where_str += " and  c.text_area2 like '%{}%'".format(self.xl_request.get("TextArea2"))
        if self.xl_request.get("TextArea3"):
            where_str += " and  c.text_area3 like '%{}%'".format(self.xl_request.get("TextArea3"))
        if self.xl_request.get("TextArea4"):
            where_str += " and  c.text_area4 like '%{}%'".format(self.xl_request.get("TextArea4"))
        if self.xl_request.get("TextArea4"):
            where_str += " and  c.text_area4 like '%{}%'".format(self.xl_request.get("TextArea4"))

        if self.xl_request.get("DateCustomField1"):
            date1 = Object.date_converter(self.xl_request.get("DateCustomField1"))
            # print date1
            where_str += " and  c.date_custom_field1 = '%s'" %date1

        if self.xl_request.get("DateCustomField2"):
            date2 = Object.date_converter(self.xl_request.get("DateCustomField2"))
            # print date2
            where_str += " and  c.date_custom_field2 = '%s'" %date2

        if self.xl_request.get("DateCustomField3"):
            date3 = Object.date_converter(self.xl_request.get("DateCustomField3"))
            # print date3
            where_str += " and  cc.date_custom_field3 = '%s'" %date3

        if self.xl_request.get("DateCustomField4"):
            date4 = Object.date_converter(self.xl_request.get("DateCustomField4"))
            # print date4
            where_str += " and  cc.date_custom_field4 = '%s'" %date4

        if self.xl_request.get("DateCustomField5"):
            date5 = Object.date_converter(self.xl_request.get("DateCustomField5"))
            where_str += " and  cc.date_custom_field5 = '%s'" %date5

        if self.xl_request.get("EventIds"):
            where_str += " and sp.recruitevent_id =  {}".format(self.xl_request.get("EventIds")[0])

        if self.xl_request.get("TestScoreFrom") and self.xl_request.get("TestScoreTo"):
            testscore_from = self.xl_request.get("TestScoreFrom")
            testscore_to = self.xl_request.get("TestScoreTo")
            where_str += " and tu.total_score  between '%s' and '%s' " %(testscore_from,testscore_to)

        if self.xl_request.get("TestIds"):
            where_str += " and tu.test_id =  {}".format(self.xl_request.get("TestIds")[0])

        if self.xl_request.get("StatusIds"):
            where_str += " and tu.status =  {}".format(self.xl_request.get("StatusIds")[0])

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
        self.ws.write(0, 0, 'Candidates Boundary search', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, self.tot_len, self.style24)
        Object.wb_result.save(output_paths.outputpaths['candidate_search_output_sheet_1'])


print("Combined Search Script Started")
Object = ExcelData()
Object.all()
print("Completed Successfully ")
Object.success_case_01 = {}
Object.overall_status()
