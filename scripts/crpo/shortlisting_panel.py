from hpro_automation.read_excel import *
import itertools
from operator import itemgetter
import requests
import json
import time
import datetime
from hpro_automation import (login, api, input_paths, output_paths, db_login, work_book)


class SCAutomation(login.CRPOLogin, db_login.DBConnection, work_book.WorkBook):
    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(SCAutomation, self).__init__()
        self.db_connection()

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 24)))
        self.Actual_Success_case = []
        self.success_case_01 = {}

        file_path = input_paths.inputpaths['Shortlisting_Input_sheet']
        sheet_index = 0
        excel_read_obj.excel_read(file_path, sheet_index)

    def excel_headers(self):
        self.main_headers = ['Status', 'CandidateID', 'ApplicantID', 'EventID', 'JobID', 'SLCID',
                             'TestID', 'Expected Status', 'DB_actual Status']
        self.headers_with_style2 = ['DB_actual Status']
        self.headers_with_style15 = ['Status', 'CandidateID', 'ApplicantID', 'EventID', 'JobID', 'SLCID',
                                     'TestID', 'Expected Status']
        self.file_headers_col_row()

    # Reading XL data
    def applicantDataRead(self):
        # self.kk1 = []
        self.xl_applicant_id = []
        self.applicant_json_data = []
        for i in excel_read_obj.complete_excel_data:
            self.xl_applicant_id.append(int(i.get('applicantId')))
            # print self.xl_applicant_id
            local = str(i.get('testId'))
            applicant_test_id = [int(float(b)) for b in local.split(',')]

            local = str(i.get('scId'))
            applicant_scid = [int(float(b)) for b in local.split(',')]

            self.convert_json = {"applicantId": int(i.get('applicantId')), "eventId": int(i.get('eventId')),
                                 "jobId": int(i.get('jobId')), "mjrId": int(i.get('mjrId')),
                                 "mjrStatusId": int(i.get('mjrStatusId')), "testId": applicant_test_id,
                                 "scId": applicant_scid}
            self.applicant_json_data.append(self.convert_json)
            # print self.applicant_json_data
            # print type(self.applicant_json_data)
        self.totalapplicantCount = len(self.xl_applicant_id)
        a = self.xl_applicant_id
        self.xl_all_applicant_id = ','.join(str(v) for v in a)
        print(self.xl_all_applicant_id)
        print(self.totalapplicantCount)

    def groupby_MJR_TEST_SLC(self):

        self.lambda_function('ChangeApplicant_Status')
        self.headers['APP-NAME'] = 'crpo'

        # Sort applicant data by `mjrid, Testid,scid,JobId and Eventid` key.
        self.applicant_json_data = sorted(self.applicant_json_data,
                                          key=itemgetter('eventId', 'jobId', 'mjrId', 'testId', 'scId'))
        # print self.applicant_json_data
        # Display data grouped by `mjrid, Testid,scid,JobId and Eventid` key.
        for key, value in itertools.groupby(self.applicant_json_data,
                                            key=itemgetter('eventId', 'jobId', 'mjrId', 'testId', 'scId')):
            # print key
            self.all_mjr_applicants = []
            self.to_status_id = []
            for iter in self.applicant_json_data:
                self.value = iter
                if key[0] == self.value['eventId'] and key[1] == self.value['jobId'] and \
                        key[2] == self.value['mjrId'] and key[3] == self.value['testId'] and key[4] == self.value[
                    'scId']:
                    self.all_mjr_applicants.append(self.value['applicantId'])
                    self.to_status_id.append(self.value['mjrStatusId'])
            # print self.all_mjr_applicants
            # print self.to_status_id
            self.data = {"ApplicantIds": self.all_mjr_applicants,
                         "EventId": key[0],
                         "JobRoleId": key[1],
                         "ToStatusId": self.to_status_id[0],
                         "Sync": "True", "Comments": "",
                         "InitiateStaffing": False,
                         "MjrId": key[2]}
            print(self.data)
            time.sleep(3)
            r = requests.post(api.web_api['ChangeApplicant_Status'],
                              headers=self.headers, data=json.dumps(self.data, default=str), verify=False)
            print(r.headers)
            resp_dict = json.loads(r.content)
            self.status = resp_dict['status']
            if self.status == 'OK':
                print("API Status is", self.status)

                self.data = {"eventId": key[0],
                             "jobId": key[1],
                             "statusId": self.to_status_id[0],
                             "testIds": key[3],
                             "shortlistingCriteriaIds": key[4]}
                print(self.data)
                # time.sleep(3)
                r = requests.post(api.web_api['oneClickShortlist'],
                                  headers=self.get_token, data=json.dumps(self.data, default=str), verify=False)
                resp_dict = json.loads(r.content)
                # print resp_dict

    # pending is get all applicant statuss, match excel, writeexcel
    def allApplicantStatuss(self):
        try:
            self.applicant_query = "select a.id,rs.label as currentstatus from applicant_statuss a " \
                                   "left join resume_statuss rs on a.current_status_id=rs.id where a.id " \
                                   "in(%s) order by id asc;" % (self.xl_all_applicant_id)

            query = self.applicant_query
            time.sleep(2)
            self.cursor.execute(query)
            db_applicant_details = self.cursor.fetchall()
            self.db_all_applicants = []
            self.db_applicant_id = []
            self.db_applicant_status = []
            self.db_j = 0

            # converting json is important for comparision
            for new_data in db_applicant_details:
                self.db_applicant_id.append(new_data[0])
                self.db_applicant_status.append(new_data[1])
                self.convert_json1 = {'dbApplicantID': self.db_applicant_id[self.db_j],
                                      'dbApplicantStatus': self.db_applicant_status[self.db_j]}
                self.db_j += 1
                self.db_all_applicants.append(self.convert_json1)
        except:
            print("DB connection Error")

    def match_db_excel(self):
        total_db_count = len(self.db_all_applicants)
        for iteration_count in range(0, self.totalapplicantCount):
            for dbdata in range(0, total_db_count):
                dbvalue = self.db_all_applicants[dbdata]
                excel_value = excel_read_obj.complete_excel_data[iteration_count]
                if excel_value.get('applicantId') == dbvalue['dbApplicantID']:
                    if excel_value.get('expectedStatus') == dbvalue['dbApplicantStatus']:
                        self.db_mess = 'pass'
                        self.success_case_01 = 'Pass'
                        self.excel_write(excel_value.get('candidateID'),
                                         excel_value.get('applicantId'),
                                         excel_value.get('expectedStatus'),
                                         dbvalue['dbApplicantStatus'],
                                         self.db_mess,
                                         self.style14, excel_value.get('eventId'),
                                         excel_value.get('jobId'),
                                         excel_value.get('testId'),
                                         excel_value.get('scId'))
                    else:
                        self.db_mess = 'status not matched with excel and DB'
                        self.excel_write(excel_value.get('candidateID'),
                                         excel_value.get('applicantId'),
                                         excel_value.get('expectedStatus'),
                                         dbvalue['dbApplicantStatus'],
                                         self.db_mess,
                                         self.style14, excel_value.get('eventId'),
                                         excel_value.get('jobId'),
                                         excel_value.get('testId'),
                                         excel_value.get('scId'))

    def excel_write(self, candidate_id, xl_aid, xl_exp_status, db_status, db_mess, __style, event_id, job_id, testid,
                    scid):
        self.ws.write(self.rowsize, 1, candidate_id, self.style16)
        self.ws.write(self.rowsize, 2, xl_aid, self.style16)
        self.ws.write(self.rowsize, 3, event_id, self.style16)
        self.ws.write(self.rowsize, 4, job_id, self.style16)
        self.ws.write(self.rowsize, 5, str(testid), self.style16)
        self.ws.write(self.rowsize, 6, str(scid), self.style16)
        self.ws.write(self.rowsize, 7, xl_exp_status, self.style16)
        self.ws.write(self.rowsize, 8, db_status, __style)
        self.ws.write(self.rowsize, 0, db_mess, self.style26)
        self.rowsize = self.rowsize + 1
        ob.wb_Result.save(output_paths.outputpaths['SC_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)

    def overall_status(self):
        self.ws.write(0, 0, 'Shortlisting Panel', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'StartTime', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        ob.wb_Result.save(output_paths.outputpaths['SC_Output_sheet'])


ob = SCAutomation()
ob.excel_headers()
ob.applicantDataRead()
ob.groupby_MJR_TEST_SLC()
time.sleep(5)
ob.allApplicantStatuss()
ob.match_db_excel()
ob.success_case_01 = {}
ob.overall_status()
