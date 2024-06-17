from hpro_automation import (login, input_paths, output_paths, work_book)
from hpro_automation.api import *
import json
import requests
import xlrd
import datetime


class RescheduleInterview(login.CommonLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(RescheduleInterview, self).__init__()
        self.common_login('admin')
        self.crpo_app_name = self.app_name.strip()
        print(self.crpo_app_name)

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Event_id = []  # [] Initialising data from excel sheet to the variables
        self.xl_Applicant_id = []
        self.xl_Job_id = []
        self.xl_type = []
        self.xl_Schedule_Datetime = []
        self.xl_stage_id = []
        self.xl_interviewers_id = []
        self.xl_Schedule_Comment = []
        self.xl_location = []

        # --------------------------
        # Reschedule and cancel data
        # --------------------------
        self.xl_Reschedule_DateTime = []
        self.xl_Reschedule_type = []
        self.xl_Reschedule_add_interviewers = []
        self.xl_Reschedule_remove_interviewers = []
        self.xl_Reschedule_comment = []
        self.xl_Exception_message = []

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 7)))
        self.Actual_Success_case = []

        # -----------------------------------------------------------------------------------
        # Dictionaries for Interview_schedule, interview_feedback, interview_feedback_details
        # -----------------------------------------------------------------------------------
        self.ir = {}
        self.reschedule_ir = {}

        self.is_success = {}
        self.is_reschedule_success = {}

        self.message = {}
        self.reschedule_message = {}

        self.candidate_details_dict = {}
        self.applicant_details_dict = {}
        self.data = {}
        self.interviewers = {}

        self.updated_candidate_details_dict = {}
        self.updated_applicant_details_dict = {}
        self.updated_data = {}
        self.updated_interviewers = {}

        self.final_status = {}

        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}

    def excel_headers(self):
        self.main_headers = ['Comparision', 'Status', 'Interview_Request', 'Interviewers',
                             'interview_type', 'ScheduleOn', 'interviewStatus', 'Reschedule_ir_id',
                             'ReSchedule_Interviewers', 'ReSchedule_type', 'ReScheduleOn', 'Reschedule_Status',
                             'ApplicantId', 'Exception_message']
        self.headers_with_style2 = ['Comparision', 'Status', 'Interview_Request', 'Interviewers', 'interview_type',
                                    'ScheduleOn', 'interviewStatus']
        self.headers_with_style9 = ['Reschedule_ir_id', 'add_interviewers', 'Removed_interviewers', 'ReSchedule_type',
                                    'Reschedule_Status', 'ReScheduleOn', 'ReSchedule_Interviewers']
        self.file_headers_col_row()

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Reschedule_input_sheet'])
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[0] != '':
                    self.xl_Event_id.append(int(rows[0]))
                else:
                    self.xl_Event_id.append(None)

                if rows[1] is not None and rows[1] != '':
                    self.xl_Applicant_id.append(int(rows[1]))
                else:
                    self.xl_Applicant_id.append(None)

                if rows[2] is not None and rows[2] != '':
                    self.xl_Job_id.append(int(rows[2]))
                else:
                    self.xl_Job_id.append(None)

                if rows[3] is not None and rows[3] != '':
                    self.xl_type.append(int(rows[3]))
                else:
                    self.xl_type.append(None)

                if rows[4] is not None and rows[4] != '':
                    self.xl_Schedule_Datetime.append(str(rows[4]))
                else:
                    self.xl_Schedule_Datetime.append(None)

                if rows[5] is not None and rows[5] != '':
                    self.xl_stage_id.append(int(rows[5]))
                else:
                    self.xl_stage_id.append(None)

                if rows[6] is not None and rows[6] != '':
                    int_ids = list(map(int, rows[6].split(',') if isinstance(rows[6], str) else [rows[6]]))
                    self.xl_interviewers_id.append(int_ids)
                else:
                    self.xl_interviewers_id.append(None)

                if rows[7] is not None and rows[7] != '':
                    self.xl_Schedule_Comment.append(str(rows[7]))
                else:
                    self.xl_Schedule_Comment.append(None)

                if rows[8] is not None and rows[8] != '':
                    self.xl_location.append(int(rows[8]))
                else:
                    self.xl_location.append(None)

                if rows[9] is not None and rows[9] != '':
                    self.xl_Reschedule_DateTime.append(str(rows[9]))
                else:
                    self.xl_Reschedule_DateTime.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_Reschedule_type.append(int(rows[10]))
                else:
                    self.xl_Reschedule_type.append(None)

                if rows[11] is not None and rows[11] != '':
                    int_ids = list(map(int, rows[11].split(',') if isinstance(rows[11], str) else [rows[11]]))
                    self.xl_Reschedule_add_interviewers.append(int_ids)
                else:
                    self.xl_Reschedule_add_interviewers.append(None)

                if rows[12] is not None and rows[12] != '':
                    int_ids = list(map(int, rows[12].split(',') if isinstance(rows[12], str) else [rows[12]]))
                    self.xl_Reschedule_remove_interviewers.append(int_ids)
                else:
                    self.xl_Reschedule_remove_interviewers.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_Reschedule_comment.append(str(rows[13]))
                else:
                    self.xl_Reschedule_comment.append(None)

                if rows[13] is not None and rows[14] != '':
                    self.xl_Exception_message.append(str(rows[14]))
                else:
                    self.xl_Exception_message.append(None)

            print('Excel data initiated is Done')

        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):

        self.lambda_function('Schedule')
        self.headers['APP-NAME'] = self.crpo_app_name

        try:
            schedule_request = [{
                "isConsultantRound": False,
                "interviewDate": self.xl_Schedule_Datetime[loop],
                "interviewTime": "",
                "interviewType": self.xl_type[loop],
                "interviewerIds": self.xl_interviewers_id[loop],
                "jobId": self.xl_Job_id[loop],
                "stageId": self.xl_stage_id[loop],
                "locationId": self.xl_location[loop],  # API default send bangalore location
                "secondaryInterviewerIds": [],
                "recruiterComment": self.xl_Schedule_Comment[loop],
                "recruitEventId": self.xl_Event_id[loop],
                "applicantIds": [self.xl_Applicant_id[loop]]
            }]
            print(schedule_request)
            scheduling_interviews = requests.post(self.webapi, headers=self.headers,
                                                  data=json.dumps(schedule_request, default=str), verify=False)
            print(scheduling_interviews.headers)
            schedule_response = json.loads(scheduling_interviews.content)
            # print(json.dumps(schedule_response, indent=2))
            data = schedule_response['data']
            # print(json.dumps(data, indent=2))
            # print('***--------------------------------------------------------***')

            if schedule_response['status'] == 'OK':
                success = data['success']
                failure = data['failure']

                if data['success']:
                    for i in success:
                        self.ir = i['interviewRequestId']
                        print(self.ir)
                        print("Scheduled to interview")
                        self.message = i.get('message')
                        if self.xl_Applicant_id[loop] is not None:
                            self.is_success = True
                elif data['failure']:
                    for i in failure:
                        self.message = i.get('message')
                        print(self.message)
                        self.is_success = False
            else:
                print('Error occured while scheduling')
        except ValueError as Schedule_error:
            print(Schedule_error)

    def reschedule_interview(self, loop):

        self.lambda_function('Reschedule')
        self.headers['APP-NAME'] = self.crpo_app_name

        reschedule_request = [{
            "interviewRequestId": self.ir,
            "interviewType": self.xl_Reschedule_type[loop],
            "interviewers": self.xl_Reschedule_add_interviewers[loop],
            "interviewDate": self.xl_Reschedule_DateTime[loop],
            "recruiterComment": self.xl_Reschedule_comment[loop],
            "removedInterviewers": self.xl_Reschedule_remove_interviewers[loop],
            "locationId": 25085
        }]
        print(reschedule_request)

        reschedule_api = requests.post(self.webapi, headers=self.headers,
                                       data=json.dumps(reschedule_request, default=str), verify=False)
        print(reschedule_api.headers)
        reschedule_response = json.loads(reschedule_api.content)
        data = reschedule_response['data']

        if reschedule_response['status'] == 'OK':
            success = data['success']
            failure = data['failure']
            print(failure)

            if data['success']:
                for i in success:
                    self.reschedule_ir = i['interviewId']
                    print(self.reschedule_ir)
                    print("ReScheduled to interview")
                    if self.xl_Applicant_id[loop] is not None:
                        self.is_reschedule_success = True
            elif data['failure']:
                for i in failure:
                    self.reschedule_message = i.get('message')
                    print(self.message)
                    self.is_reschedule_success = False

    def interview_request_details(self):

        self.lambda_function('InterviewRequest_details')
        self.headers['APP-NAME'] = self.crpo_app_name

        ir_details_request = {
            "search": {
                "interviewrequests": [self.ir]
            }}
        ir_details_api = requests.post(self.webapi, headers=self.headers,
                                       data=json.dumps(ir_details_request, default=str), verify=False)
        print(ir_details_api.headers)
        ir_details_response = json.loads(ir_details_api.content)
        data = ir_details_response['data']
        for i in data:
            self.data = i
            self.candidate_details_dict = i.get('candidate')
            self.applicant_details_dict = i.get('applicant')

            interviewer_details_dict = i.get('interviewers')
            int1 = interviewer_details_dict.keys()
            self.interviewers = ', '.join(int1)

    def updated_interview_request_details(self):

        self.lambda_function('InterviewRequest_details')
        self.headers['APP-NAME'] = self.crpo_app_name

        ir_details_request = {
            "search": {
                "interviewrequests": [self.ir]
            }}
        ir_details_api = requests.post(self.webapi, headers=self.headers,
                                       data=json.dumps(ir_details_request, default=str), verify=False)
        print(ir_details_api.headers)
        ir_details_response = json.loads(ir_details_api.content)
        data = ir_details_response['data']
        for i in data:
            self.updated_data = i
            self.updated_candidate_details_dict = i.get('candidate')
            self.updated_applicant_details_dict = i.get('applicant')

            updated_interviewer_details_dict = i.get('interviewers')
            updated_int = updated_interviewer_details_dict.keys()
            self.updated_interviewers = ', '.join(updated_int)
            self.final_status = True

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 3, str(self.xl_interviewers_id[loop]))
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_type[loop] == 2:
            self.ws.write(self.rowsize, 4, 'In_Person')
        elif self.xl_type[loop] == 3:
            self.ws.write(self.rowsize, 4, 'Hirepro_Video')
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 5, self.xl_Schedule_Datetime[loop])
        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 8,
                      str(self.xl_Reschedule_add_interviewers[loop])
                      if self.xl_Reschedule_add_interviewers[loop] else 'Empty')
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_Reschedule_type[loop] == 2:
            self.ws.write(self.rowsize, 9, 'In_Person')
        elif self.xl_Reschedule_type[loop] == 3:
            self.ws.write(self.rowsize, 9, 'Hirepro_Video')
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 10, self.xl_Reschedule_DateTime[loop])
        self.ws.write(self.rowsize, 12, self.xl_Applicant_id[loop])
        self.ws.write(self.rowsize, 13, self.xl_Exception_message[loop])

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.final_status:
            self.ws.write(self.rowsize, 1, 'Pass', self.style26)
            self.success_case_01 = 'Pass'

        elif self.xl_Exception_message[loop]:
            if self.message:
                if self.xl_Exception_message[loop] and 'already scheduled for interview' in self.message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_02 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)

            elif self.reschedule_message:
                if self.xl_Exception_message[loop] and 'Unable to reschedule' in self.reschedule_message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_03 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if self.ir:
            self.ws.write(self.rowsize, 2, self.ir, self.style26)
        else:
            self.ws.write(self.rowsize, 2, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.interviewers:
            if self.xl_interviewers_id[loop] is None:
                self.ws.write(self.rowsize, 3, 'NoInterviewers', self.style7)
            else:
                self.ws.write(self.rowsize, 3, self.interviewers, self.style8)
        else:
            self.ws.write(self.rowsize, 3, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.data and self.data.get('typeOfInterview'):
            if self.xl_type[loop] == self.data.get('typeOfInterviewEnum'):
                if self.xl_type[loop] is None:
                    self.ws.write(self.rowsize, 4, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 4, self.data.get('typeOfInterview'), self.style8)
            else:
                self.ws.write(self.rowsize, 4, None)
        else:
            self.ws.write(self.rowsize, 4, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.data and self.data.get('scheduledOn'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 5, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 5, self.data.get('scheduledOn'), self.style8)
            else:
                self.ws.write(self.rowsize, 5, None)
        else:
            self.ws.write(self.rowsize, 5, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_details_dict and self.applicant_details_dict.get('currentStatus'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 6, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 6, self.applicant_details_dict.get('currentStatus'), self.style8)
            else:
                self.ws.write(self.rowsize, 6, None)
        else:
            self.ws.write(self.rowsize, 6, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.reschedule_ir:
            self.ws.write(self.rowsize, 7, self.reschedule_ir, self.style26)
        else:
            self.ws.write(self.rowsize, 7, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_interviewers:
            if self.xl_Reschedule_add_interviewers is None:
                self.ws.write(self.rowsize, 8, 'NoInterviewers', self.style7)
            else:
                self.ws.write(self.rowsize, 8, self.updated_interviewers, self.style8)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_data and self.updated_data.get('typeOfInterview'):
            if self.xl_Reschedule_type[loop] == self.updated_data.get('typeOfInterviewEnum'):
                if self.xl_Reschedule_type[loop] is None:
                    self.ws.write(self.rowsize, 9, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 9, self.updated_data.get('typeOfInterview'), self.style8)
            else:
                self.ws.write(self.rowsize, 9, None)
        else:
            self.ws.write(self.rowsize, 9, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_data and self.updated_data.get('scheduledOn'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 10, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 10, self.updated_data.get('scheduledOn'), self.style8)
            else:
                self.ws.write(self.rowsize, 10, None)
        else:
            self.ws.write(self.rowsize, 10, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.updated_applicant_details_dict and self.updated_applicant_details_dict.get('currentStatus'):
            if self.xl_Applicant_id[loop] == self.updated_applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 11, self.updated_applicant_details_dict.get('currentStatus'),
                                  self.style8)
            else:
                self.ws.write(self.rowsize, 11, None)
        else:
            self.ws.write(self.rowsize, 11, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_details_dict and self.applicant_details_dict.get('applicantId'):
            if self.xl_Applicant_id[loop] == self.applicant_details_dict.get('applicantId'):
                if self.xl_Applicant_id is None:
                    self.ws.write(self.rowsize, 12, 'Empty_value', self.style8)
                else:
                    self.ws.write(self.rowsize, 12, self.applicant_details_dict.get('applicantId'), self.style8)
            else:
                self.ws.write(self.rowsize, 12, self.xl_Applicant_id[loop], self.style8)
        else:
            self.ws.write(self.rowsize, 12, None)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Exception_message[loop]:
            if self.message is not None:
                if 'Unable to' in self.message:
                    self.ws.write(self.rowsize, 13, self.message, self.style8)
                elif 'Applicant has' in self.message:
                    self.ws.write(self.rowsize, 13, self.message, self.style8)
                else:
                    self.ws.write(self.rowsize, 13, self.message, self.style3)
            elif self.reschedule_message:
                self.ws.write(self.rowsize, 13, self.reschedule_message, self.style8)
            else:
                self.ws.write(self.rowsize, 13, self.message, self.style3)
        elif self.reschedule_message:
            self.ws.write(self.rowsize, 13, self.reschedule_message, self.style3)
        else:
            self.ws.write(self.rowsize, 13, self.message, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Reschdeule_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)

    def overall_status(self):
        self.ws.write(0, 0, 'Reschedule', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Login Server', self.style23)
        self.ws.write(0, 3, login_server, self.style24)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        self.ws.write(0, 6, 'APP Name', self.style23)
        self.ws.write(0, 7, self.crpo_app_name, self.style24)
        self.ws.write(0, 8, 'No.of Test cases', self.style23)
        self.ws.write(0, 9, Total_count, self.style24)
        self.ws.write(0, 10, 'Start Time', self.style23)
        self.ws.write(0, 11, self.start_time, self.style26)
        Object.wb_Result.save(output_paths.outputpaths['Reschdeule_Output_sheet'])


Object = RescheduleInterview()
Object.excel_headers()
Object.excel_data()
Total_count = len(Object.xl_Event_id)
print("Number of Rows ::", Total_count)

try:
    if Object.login == 'OK':
        for looping in range(0, Total_count):
            print("Iteration Count is ::", looping)
            Object.schedule_interview(looping)
            if Object.is_success:
                Object.interview_request_details()
                Object.reschedule_interview(looping)
                if Object.is_reschedule_success:
                    Object.updated_interview_request_details()
            Object.output_excel(looping)
            print('Excel data is ready')

            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.is_success = {}
            Object.is_reschedule_success = {}
            Object.ir = {}
            Object.reschedule_ir = {}
            Object.message = {}
            Object.reschedule_message = {}

            Object.candidate_details_dict = {}
            Object.applicant_details_dict = {}
            Object.interviewers = {}

            Object.updated_candidate_details_dict = {}
            Object.updated_applicant_details_dict = {}
            Object.updated_interviewers = {}

            Object.data = {}
            Object.updated_data = {}

            Object.final_status = {}
            Object.success_case_01 = {}
            Object.success_case_02 = {}
            Object.success_case_03 = {}
    Object.overall_status()

except AttributeError as Object_error:
    print(Object_error)
