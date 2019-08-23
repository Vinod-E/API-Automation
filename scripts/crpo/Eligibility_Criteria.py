import requests
import json
import xlrd
import datetime
import time
from hpro_automation import (login, input_paths, output_paths, db_login, work_book)


class ECAutomation(login.CommonLogin, db_login.DBConnection, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(ECAutomation, self).__init__()
        self.common_login('crpo')
        self.db_connection('amsin')

        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        # excel values are assigned to local variable
        self.xl_candidate_id = []
        self.xl_applicant_id = []
        self.xl_expected_result = []
        self.xl_ec_id = []
        self.xl_event_Id = []
        self.xl_job_Id = []
        self.xl_job_Id = []
        self.xl_mjr_Id = []
        self.xl_to_status__Id = []
        self.xl_positive_status__Id = []
        self.xl_negative_status__Id = []
        self.xl_ec_configuration__Id = []
        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 103)))
        self.Actual_Success_case = []
        self.success_case_01 = {}

        self.excel_headers()

    def excel_headers(self):
        self.main_headers = ['Actual_status', 'Event Id', 'Job Id', 'CandidateID', 'Excel ApplicantId',
                             'DB Applicant ID', 'EC ID', 'Expected Status', 'DB Status']
        self.headers_with_style2 = ['DB Status', 'Actual_status']
        self.file_headers_col_row()
    # -------------------------------------------
    # Reading Input data from excel
    # -------------------------------------------
    def Data_read(self):
        wb = xlrd.open_workbook(input_paths.inputpaths['EC_Input_sheet'])
        sh1 = wb.sheet_by_index(0)
        i = 1
        for i in range(1, sh1.nrows):
            rownum = i
            rows = sh1.row_values(rownum)
            self.xl_candidate_id.append(int(rows[0]))
            self.xl_applicant_id.append(int(rows[1]))
            self.xl_expected_result.append(rows[2])
            self.xl_ec_id.append(int(rows[3]))
            self.xl_event_Id.append(int(rows[4]))
            self.xl_job_Id.append(int(rows[5]))
            self.xl_mjr_Id.append(int(rows[6]))
            self.xl_to_status__Id.append(int(rows[7]))
            self.xl_positive_status__Id.append(int(rows[8]))
            self.xl_negative_status__Id.append(int(rows[9]))
            self.xl_ec_configuration__Id.append(int(rows[10]))
            self.applicant_id = self.xl_applicant_id
            self.candidate_id = self.xl_candidate_id
            self.expected_results = self.xl_expected_result
            self.ec_id = self.xl_ec_id

    # ------------------------------------------------------------------------------------------------------------------
    # Fetching applicant status data from DB
    # below method is used for Database connectivity and query execution
    # conn.close is important for every iteration otherwise, it will wrong data from local image.
    # ------------------------------------------------------------------------------------------------------------------
    def Fetch_Applicants_DB(self, iteration_count):
        try:

            self.applicant_query = "select  ap.candidate_id,ap.id as applicant_id,ap.current_status_id,rs.label as status " \
                                   "from applicant_statuss ap left join resume_statuss rs on ap.current_status_id = rs.id " \
                                   "where ap.id ='%s';" % (self.applicant_id[iteration_count])
            query = self.applicant_query
            time.sleep(2)
            self.cursor.execute(query)
            data_to_return = self.cursor.fetchall()

            alldata = data_to_return
            tot = len(alldata)
            if tot > 0:
                for item in range(0, tot):
                    self.data1 = alldata[item]
                    self.db_candidate_id = self.data1[0]
                    self.db_applicant_id = self.data1[1]
                    self.db_status_id = self.data1[2]
                    self.db_status = self.data1[3]
            else:
                print("No Data From Database")
        except:
            print("DB connection Error")

    def api_main(self, iteration_count):
        # -------------------------------------------
        # Updating EC configuration at event level
        # -------------------------------------------
        self.lambda_function('createOrUpdateEcConfig')
        self.headers['APP-NAME'] = 'crpo'

        self.data = {"ecConfigurations": [{"id": self.xl_ec_configuration__Id[iteration_count],
                                           "jobRoleId": self.xl_job_Id[iteration_count],
                                           "eventId": self.xl_event_Id[iteration_count],
                                           "ecId": self.ec_id[iteration_count],
                                           "positiveStatusId": self.xl_positive_status__Id[iteration_count],
                                           "negativeStatusId": self.xl_negative_status__Id[iteration_count]}]}
        r = requests.post(self.webapi, headers=self.headers,
                          data=json.dumps(self.data, default=str), verify=False)
        print(r.headers)
        time.sleep(1)
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        if self.status == 'OK':
            print("EC Updated Successfully")
            print("EC Status Code is", self.status)
            # ---------------------------
            # Changing applicant status
            # ---------------------------
            self.lambda_function('ChangeApplicant_Status')
            self.headers['APP-NAME'] = 'crpo'

            self.data = {"ApplicantIds": [self.applicant_id[iteration_count]],
                         "EventId": self.xl_event_Id[iteration_count],
                         "JobRoleId": self.xl_job_Id[iteration_count],
                         "ToStatusId": self.xl_to_status__Id[iteration_count],
                         "Sync": "True",
                         "Comments": "",
                         "InitiateStaffing": False,
                         "MjrId": self.xl_mjr_Id[iteration_count]}
            r = requests.post(self.webapi, headers=self.headers,
                              data=json.dumps(self.data, default=str), verify=False)
            print(r.headers)
            resp_dict = json.loads(r.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                print("Status Change API is Executed")
                print("Status is", self.status)
                time.sleep(2)
            else:
                print("Status Change API is not executed")
                print("Status is", self.status)
        else:
            print("EC Not Updated Successfully")
            print("EC Status Code is", self.status)

    # -------------------------------------
    # Compairing data from Excel to DB
    # -------------------------------------
    def match_db_excel(self, iteration_count):
        if self.applicant_id[iteration_count] == self.db_applicant_id:
            if self.expected_results[iteration_count] == self.db_status:
                print("DB - Status Matched With Expected Status")
                self.db_mess = 'pass'
                self.success_case_01 = 'Pass'
                self.excel_write(self.candidate_id[iteration_count],
                                 self.applicant_id[iteration_count],
                                 self.db_applicant_id,
                                 self.expected_results[iteration_count],
                                 self.db_status,
                                 self.db_mess,
                                 self.style14,
                                 self.xl_event_Id[iteration_count],
                                 self.xl_job_Id[iteration_count],
                                 self.ec_id[iteration_count])
            else:
                print("DB - Status Not Matched  With Expected Status")
                self.db_mess = 'status not matched with excel and DB'
                self.excel_write(self.candidate_id[iteration_count],
                                 self.applicant_id[iteration_count],
                                 self.db_applicant_id,
                                 self.expected_results[iteration_count],
                                 self.db_status,
                                 self.db_mess,
                                 self.style13,
                                 self.xl_event_Id[iteration_count],
                                 self.xl_job_Id[iteration_count],
                                 self.ec_id[iteration_count])
        else:
            pass
            self.db_mess = 'Excel Applicant id not matched with DB applicant'
            self.excel_write(self.candidate_id[iteration_count],
                             self.applicant_id[iteration_count],
                             self.db_applicant_id,
                             self.expected_results[iteration_count],
                             self.db_status,
                             self.db_mess,
                             self.style13,
                             self.xl_event_Id[iteration_count],
                             self.xl_job_Id[iteration_count],
                             self.ec_id[iteration_count])

    # ------------------------
    # Writing output to excel
    # ------------------------
    def excel_write(self, candidate_id, ex_aid, db_aid, ex_exp_status, db_status, db_mess, __style, event_id, job_id,
                    ec_id):
        self.ws.write(self.rowsize, 1, event_id, self.style12)
        self.ws.write(self.rowsize, 2, job_id, self.style12)
        self.ws.write(self.rowsize, 3, candidate_id, self.style12)
        self.ws.write(self.rowsize, 4, ex_aid, self.style12)
        self.ws.write(self.rowsize, 5, db_aid, self.style12)
        self.ws.write(self.rowsize, 7, ex_exp_status, self.style12)
        self.ws.write(self.rowsize, 8, db_status, self.style14)
        self.ws.write(self.rowsize, 6, ec_id, self.style12)
        self.ws.write(self.rowsize, 0, db_mess, self.style26)
        self.rowsize = self.rowsize + 1
        ob.wb_Result.save(output_paths.outputpaths['EC_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Eligibility Criteria', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'No.of Test cases', self.style23)
        self.ws.write(0, 7, tot_count, self.style24)
        ob.wb_Result.save(output_paths.outputpaths['EC_Output_sheet'])


ob = ECAutomation()
ob.Data_read()
tot_count = len(ob.applicant_id)
print(tot_count)
if ob.login == 'OK':
    for iteration_count in range(0, tot_count):
        print("Iteration Count is '%d'" % iteration_count)
        ob.api_main(iteration_count)
        ob.Fetch_Applicants_DB(iteration_count)
        ob.match_db_excel(iteration_count)
        ob.success_case_01 = {}
ob.overall_status()
