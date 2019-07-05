from hpro_automation import (login, input_paths, output_paths, work_book)
import json
import requests
import xlrd
import datetime


class InterviewFeedback(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(InterviewFeedback, self).__init__()

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Event_id = []  # [] Initialising data from excel sheet to the variables
        self.xl_Applicant_id = []
        self.xl_Job_id = []
        self.xl_type = []
        self.xl_Datetime = []
        self.xl_stage_id = []
        self.xl_interviewers_id = []
        self.xl_Schedule_Comment = []
        self.xl_location = []

        self.xl_Skill_id_01 = []
        self.xl_Skill_score_01 = []
        self.xl_Skill_id_02 = []
        self.xl_Skill_score_02 = []
        self.xl_Skill_id_03 = []
        self.xl_Skill_score_03 = []
        self.xl_Skill_id_04 = []
        self.xl_Skill_score_04 = []
        self.xl_skill_comment = []
        self.xl_decision = []
        self.xl_duration = []
        self.xl_int_datetime = []
        self.xl_Over_all_comment = []
        self.xl_partial_feedback = []
        self.xl_exception_message = []

        # -----------------------------------------
        # Update details / Partial feedback details
        # -----------------------------------------

        self.xl_updated_duration = []
        self.xl_Updated_Over_all_comment = []
        self.xl_update_Skill_comment = []
        self.xl_update_skill_score_01 = []
        self.xl_update_stage = []

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 33)))
        self.Actual_Success_case = []

        # -----------------------------------------------------------------------------------
        # Dictionaries for Interview_schedule, interview_feedback, interview_feedback_details
        # -----------------------------------------------------------------------------------
        self.ir = {}
        self.i_r = self.ir
        self.is_success = {}
        self.is_s = self.is_success
        self.is_feedback = {}
        self.i_f = self.is_feedback
        self.message = {}
        self.m = self.message

        self.feedback = {}
        self.f = self.feedback
        self.feedback_data = {}
        self.fd = self.feedback_data

        self.updated_feedback_data = {}
        self.u_fd = self.updated_feedback_data
        self.updated_feedback = {}
        self.u_feed = self.updated_feedback

        self.partial_data = {}
        self.pd = self.partial_data
        self.feedback_message = {}
        self.fm = self.feedback_message
        self.applicant_details = {}
        self.ad = self.applicant_details

        # -------------------
        # Skill dictionaries
        # -------------------
        self.skill_dict_1 = {}
        self.skill_1 = self.skill_dict_1
        self.skill_dict_2 = {}
        self.skill_2 = self.skill_dict_2
        self.skill_dict_3 = {}
        self.skill_3 = self.skill_dict_3
        self.skill_dict_4 = {}
        self.skill_4 = self.skill_dict_4

        self.filledFeedbackDetails = {}
        self.ffd = self.filledFeedbackDetails
        self.skillAssessed_details = {}
        self.sad = self.skillAssessed_details

        # ---------------------------
        # Skill updated dictionaries
        # ---------------------------
        self.updated_skill_dict_1 = {}
        self.updated_skill_1 = self.updated_skill_dict_1
        self.updated_skill_dict_2 = {}
        self.updated_skill_2 = self.updated_skill_dict_2
        self.updated_skill_dict_3 = {}
        self.updated_skill_3 = self.updated_skill_dict_3
        self.updated_skill_dict_4 = {}
        self.updated_skill_4 = self.updated_skill_dict_4

        self.updated_filledFeedbackDetails = {}
        self.updated_ffd = self.updated_filledFeedbackDetails
        self.updated_skillAssessed_details = {}
        self.u_sad = self.updated_skillAssessed_details
        self.decision_update = {}
        self.dr = self.decision_update
        self.updatedecision = {}
        self.ud = self.updatedecision
        self.decision_error = {}
        self.de = self.decision_error
        self.decision_updated_feedback = {}
        self.duf = self.decision_updated_feedback

        # ----------------------------
        # Partial/updated Dictionaries
        # ----------------------------
        self.pf = {}
        self.p_f = self.pf

        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.success_case_04 = {}
        self.success_case_05 = {}
        self.success_case_06 = {}

    def excel_headers(self):
        self.main_headers = ['Comparison', 'Status', 'ApplicantID', 'IR_id', 'Feedback_Message',
                             'Interviewer_Decision', 'Partial_Feedback', 'Partial_Feedback_message',
                             'Update_decision_message', 'Updated_decision', 'Scheduled_date', 'Interviewed_date',
                             'Skill_01', 'Score_01', 'Skill_02', 'Score_02', 'Skill_03', 'Score_03', 'Skill_04',
                             'Score_04', 'Duration', 'Skill_comment', 'OverAllComment', 'Update_Score_01',
                             'Update_Duration', 'Update_Skill_comment', 'Updated_OverAllComment', 'Exception_Message']
        self.headers_with_style2 = ['Comparison', 'Status', 'ApplicantID', 'IR_id', 'Feedback_Message',
                                    'Interviewer_Decision', 'Partial_Feedback', 'Partial_Feedback_message',
                                    'Update_decision_message', 'Updated_decision']
        self.headers_with_style9 = ['Scheduled_date', 'Interviewed_date', 'Update_Skill_comment',
                                    'Updated_OverAllComment', 'Update_Duration', 'Update_Score_01']
        self.file_headers_col_row()

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook(input_paths.inputpaths['Provide_feedback_Input_sheet'])
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
                    self.xl_Datetime.append(str(rows[4]))
                else:
                    self.xl_Datetime.append(None)

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
                    self.xl_Skill_id_01.append(int(rows[9]))
                else:
                    self.xl_Skill_id_01.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_Skill_score_01.append(int(rows[10]))
                else:
                    self.xl_Skill_score_01.append(None)

                if rows[11] is not None and rows[11] != '':
                    self.xl_Skill_id_02.append(int(rows[11]))
                else:
                    self.xl_Skill_id_02.append(None)

                if rows[12] is not None and rows[12] != '':
                    self.xl_Skill_score_02.append(int(rows[12]))
                else:
                    self.xl_Skill_score_02.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_Skill_id_03.append(int(rows[13]))
                else:
                    self.xl_Skill_id_03.append(None)

                if rows[14] is not None and rows[14] != '':
                    self.xl_Skill_score_03.append(int(rows[14]))
                else:
                    self.xl_Skill_score_03.append(None)

                if rows[15] is not None and rows[15] != '':
                    self.xl_Skill_id_04.append(int(rows[15]))
                else:
                    self.xl_Skill_id_04.append(None)

                if rows[16] is not None and rows[16] != '':
                    self.xl_Skill_score_04.append(int(rows[16]))
                else:
                    self.xl_Skill_score_04.append(None)

                if rows[17] is not None and rows[17] != '':
                    self.xl_skill_comment.append(str(rows[17]))
                else:
                    self.xl_skill_comment.append(None)

                if rows[18] is not None and rows[18] != '':
                    self.xl_decision.append(int(rows[18]))
                else:
                    self.xl_decision.append(None)

                if rows[19] is not None and rows[19] != '':
                    self.xl_duration.append(int(rows[19]))
                else:
                    self.xl_duration.append(None)

                if rows[20] is not None and rows[20] != '':
                    self.xl_int_datetime.append(str(rows[20]))
                else:
                    self.xl_int_datetime.append(None)

                if rows[21] is not None and rows[21] != '':
                    self.xl_Over_all_comment.append(str(rows[21]))
                else:
                    self.xl_Over_all_comment.append(None)

                if rows[22] is not None and rows[22] != '':
                    self.xl_partial_feedback.append(int(rows[22]))
                else:
                    self.xl_partial_feedback.append(None)

                if rows[23] is not None and rows[23] != '':
                    self.xl_updated_duration.append(int(rows[23]))
                else:
                    self.xl_updated_duration.append(None)

                if rows[24] is not None and rows[24] != '':
                    self.xl_Updated_Over_all_comment.append(str(rows[24]))
                else:
                    self.xl_Updated_Over_all_comment.append(None)

                if rows[25] is not None and rows[25] != '':
                    self.xl_update_Skill_comment.append(str(rows[25]))
                else:
                    self.xl_update_Skill_comment.append(None)

                if rows[26] is not None and rows[26] != '':
                    self.xl_update_skill_score_01.append(int(rows[26]))
                else:
                    self.xl_update_skill_score_01.append(None)

                if rows[27] is not None and rows[27] != '':
                    self.xl_update_stage.append(int(rows[27]))
                else:
                    self.xl_update_stage.append(None)

                if rows[28] is not None and rows[28] != '':
                    self.xl_exception_message.append(rows[28])
                else:
                    self.xl_exception_message.append(None)

            print('Excel data initiated is Done')

        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):

        self.lambda_function('Schedule')
        self.headers['APP-NAME'] = 'crpo'

        try:
            schedule_request = [{
                "isConsultantRound": False,
                "interviewDate": self.xl_Datetime[loop],
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
            scheduling_interviews = requests.post(self.webapi, headers=self.headers,
                                                  data=json.dumps(schedule_request, default=str), verify=False)
            print(scheduling_interviews.headers)
            schedule_response = json.loads(scheduling_interviews.content)
            # print (json.dumps(schedule_response, indent=2))
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

    def provide_feedback(self, loop):

        self.lambda_function('givefeedback')
        self.headers['APP-NAME'] = 'crpo'

        if self.xl_int_datetime[loop]:
            if self.xl_partial_feedback[loop] == 1:
                self.pf = True
            else:
                self.pf = False
            try:
                feedback_request = {
                    "interviewRequestId": self.ir,
                    "interviewerFeedback": [{
                        "partial_feedback": self.pf,
                        "skillsAssessed": [{
                            "skillId": self.xl_Skill_id_01[loop],
                            "skillScore": self.xl_Skill_score_01[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_02[loop],
                            "skillScore": self.xl_Skill_score_02[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_03[loop],
                            "skillScore": self.xl_Skill_score_03[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }, {
                            "skillId": self.xl_Skill_id_04[loop],
                            "skillScore": self.xl_Skill_score_04[loop],
                            "skillComment": self.xl_skill_comment[loop]
                        }],
                        "interviwerIds": self.xl_interviewers_id[loop],
                        "applicantId": self.xl_Applicant_id[loop],
                        "interviewerDecision": self.xl_decision[loop],
                        "interviewerComment": self.xl_Over_all_comment[loop],
                        "interviewDuration": self.xl_duration[loop],
                        "interviewedDate": self.xl_int_datetime[loop]
                    }]
                }
                providing_feedback = requests.post(self.webapi, headers=self.headers,
                                                   data=json.dumps(feedback_request, default=str), verify=False)
                print(providing_feedback.headers)
                feedback_response = json.loads(providing_feedback.content)
                # print (json.dumps(feedback_response, indent=2))
                data = feedback_response.get('data')
                self.feedback_message = data.get('message')
                self.is_feedback = True
                print('Provide Feedback is Done')

                if self.xl_decision[loop] == 167097:
                    self.updatedecision = True

            except ValueError as feedback_error:
                print(feedback_error)

    def feedback_details(self, loop):

        self.lambda_function('Interview_details')
        self.headers['APP-NAME'] = 'crpo'

        try:
            details_url = requests.get(self.webapi.format(self.ir), headers=self.headers)
            print(details_url.headers)
            details_response = json.loads(details_url.content)
            # print(json.dumps(details_response, indent=2))
            # print('***--------------------------------------------------------***')
            self.feedback_data = details_response['data']
            self.filledFeedbackDetails = self.feedback_data['filledFeedbackDetails']
            applicant = self.feedback_data['applicants']
            for applicants in applicant:
                self.applicant_details = applicants

            for feedback in self.filledFeedbackDetails:
                self.feedback = feedback
                for skillAssessed_details in feedback['skillAssessed']:
                    self.skillAssessed_details = skillAssessed_details

                    if self.xl_Skill_id_01[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_1 = skillAssessed_details

                    if self.xl_Skill_id_02[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_2 = skillAssessed_details

                    if self.xl_Skill_id_03[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_3 = skillAssessed_details

                    if self.xl_Skill_id_04[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_4 = skillAssessed_details

            print('Feedback details are fetched Successfully')
        except ValueError as details_error:
            print(details_error)

    def updated_feedback_details(self, loop):

        self.lambda_function('Interview_details')
        self.headers['APP-NAME'] = 'crpo'

        try:
            details_url = requests.get(self.webapi.format(self.ir), headers=self.headers)
            print(details_url.headers)
            details_response = json.loads(details_url.content)
            # print(json.dumps(details_response, indent=2))
            # print('***--------------------------------------------------------***')
            self.updated_feedback_data = details_response['data']
            self.updated_filledFeedbackDetails = self.updated_feedback_data['filledFeedbackDetails']

            for updated_feedback in self.updated_filledFeedbackDetails:
                self.updated_feedback = updated_feedback
                for updated_skillAssessed_details in updated_feedback['skillAssessed']:
                    self.updated_skillAssessed_details = updated_skillAssessed_details

                    if self.xl_Skill_id_01[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_1 = updated_skillAssessed_details

                    if self.xl_Skill_id_02[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_2 = updated_skillAssessed_details

                    if self.xl_Skill_id_03[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_3 = updated_skillAssessed_details

                    if self.xl_Skill_id_04[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_4 = updated_skillAssessed_details

            print('Updated Feedback details are fetched Successfully')
        except ValueError as details_error:
            print(details_error)

    def update_decision(self, loop):

        self.lambda_function('updateinterviewerdecision')
        self.headers['APP-NAME'] = 'crpo'

        if self.xl_update_stage[loop]:
            update_decision_request = {
                "interviewRequestId": self.ir,
                "decisionId": self.xl_update_stage[loop]
            }
            update_decision_url = requests.post(self.webapi, headers=self.headers,
                                                data=json.dumps(update_decision_request, default=str), verify=False)
            print(update_decision_url.headers)
            update_decision_response = json.loads(update_decision_url.content)
            decision_response = update_decision_response.get('data')
            decision_error = update_decision_response.get('error')
            self.decision_update = decision_response
            self.decision_error = decision_error

    def decision_updated_feedback_details(self):

        self.lambda_function('Interview_details')
        self.headers['APP-NAME'] = 'crpo'

        try:
            details_url = requests.get(self.webapi.format(self.ir), headers=self.headers)
            print(details_url.headers)
            details_response = json.loads(details_url.content)
            # print(json.dumps(details_response, indent=2))
            # print('***--------------------------------------------------------***')
            decision_updated_feedback_data = details_response['data']
            decision_updated_filledfeedbackdetails = decision_updated_feedback_data['filledFeedbackDetails']

            for decision_updated_feedback in decision_updated_filledfeedbackdetails:
                self.decision_updated_feedback = decision_updated_feedback

        except ValueError as decision:
            print(decision)

    def partial_feedback(self, loop):

        self.lambda_function('updateinterviewerfeedback')
        self.headers['APP-NAME'] = 'crpo'

        if self.feedback['partialFeedback'] == 1:

            try:
                update_feedback = {
                    "InterviewRequestId": self.ir,
                    "FilledFormId": self.skillAssessed_details['interviewfilledfeedbackformId'],
                    "Duration": self.xl_updated_duration[loop],
                    "OverAllComments": self.xl_Updated_Over_all_comment[loop],
                    "Skills": [{
                        "Id": self.skill_dict_1['id'],
                        "SkillId": self.xl_Skill_id_01[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_update_skill_score_01[loop]
                    }, {
                        "Id": self.skill_dict_2['id'],
                        "SkillId": self.xl_Skill_id_02[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_02[loop]
                    }, {
                        "Id": self.skill_dict_3['id'],
                        "SkillId": self.xl_Skill_id_03[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_03[loop]
                    }, {
                        "Id": self.skill_dict_4['id'],
                        "SkillId": self.xl_Skill_id_04[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_04[loop]
                    }]
                }
                partial_url = requests.post(self.webapi, headers=self.headers,
                                            data=json.dumps(update_feedback, default=str), verify=False)
                print(partial_url.headers)

                partial_response = json.loads(partial_url.content)
                self.partial_data = partial_response['data']

            except ValueError as Partial_update_error:
                print(Partial_update_error)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        self.ws.write(self.rowsize, 10, self.xl_Datetime[loop] if self.xl_Datetime[loop] else 'Empty')
        self.ws.write(self.rowsize, 11, self.xl_int_datetime[loop] if self.xl_int_datetime[loop] else 'Empty')
        self.ws.write(self.rowsize, 12, self.xl_Skill_id_01[loop] if self.xl_Skill_id_01[loop] else 'Empty')
        self.ws.write(self.rowsize, 13, self.xl_Skill_score_01[loop] if self.xl_Skill_score_01[loop] else 'Empty')
        self.ws.write(self.rowsize, 14, self.xl_Skill_id_02[loop] if self.xl_Skill_id_02[loop] else 'Empty')
        self.ws.write(self.rowsize, 15, self.xl_Skill_score_02[loop] if self.xl_Skill_score_02[loop] else 'Empty')
        self.ws.write(self.rowsize, 16, self.xl_Skill_id_03[loop] if self.xl_Skill_id_03[loop] else 'Empty')
        self.ws.write(self.rowsize, 17, self.xl_Skill_score_03[loop] if self.xl_Skill_score_03[loop] else 'Empty')
        self.ws.write(self.rowsize, 18, self.xl_Skill_id_04[loop] if self.xl_Skill_id_04[loop] else 'Empty')
        self.ws.write(self.rowsize, 19, self.xl_Skill_score_04[loop] if self.xl_Skill_score_04[loop] else 'Empty')
        self.ws.write(self.rowsize, 20, self.xl_duration[loop] if self.xl_duration[loop] else 'Empty')
        self.ws.write(self.rowsize, 21, self.xl_skill_comment[loop] if self.xl_skill_comment[loop] else 'Empty')
        self.ws.write(self.rowsize, 22, self.xl_Over_all_comment[loop] if self.xl_Over_all_comment[loop] else 'Empty')

        self.ws.write(self.rowsize, 23,
                      self.xl_update_skill_score_01[loop] if self.xl_update_skill_score_01[loop] else 'Empty')
        self.ws.write(self.rowsize, 24, self.xl_updated_duration[loop] if self.xl_updated_duration[loop] else 'Empty')
        self.ws.write(self.rowsize, 25,
                      self.xl_update_Skill_comment[loop] if self.xl_update_Skill_comment[loop] else 'Empty')
        self.ws.write(self.rowsize, 26,
                      self.xl_Updated_Over_all_comment[loop] if self.xl_Updated_Over_all_comment[loop] else 'Empty')
        self.ws.write(self.rowsize, 27, self.xl_exception_message[loop])

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.applicant_details:
            if self.decision_error and self.decision_error.get('errorDescription'):
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            elif self.applicant_details and self.applicant_details.get('applicantId'):
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_01 = 'Pass'
        elif self.message:
            if self.xl_exception_message[loop] is not None:
                if self.xl_exception_message[loop] and 'already scheduled for interview' in self.message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_02 = 'Pass'
                elif self.xl_exception_message[loop] and 'updated interviewer decision' in self.message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_03 = 'Pass'
                elif self.xl_exception_message[loop] and 'occurred while scheduling the interview' in self.message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_04 = 'Pass'
                elif self.xl_exception_message[loop] and 'configured for stage id : 167096' in self.message:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_05 = 'Pass'
            else:
                self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        elif self.message is None:
            if self.xl_exception_message[loop] is not None:
                if 'No message' in self.xl_exception_message[loop]:
                    self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                    self.success_case_06 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)

        # --------------------------------------------------------------------------------------------------------------

        self.ws.write(self.rowsize, 2, self.applicant_details.get('applicantId'), self.style12)
        # --------------------------------------------------------------------------------------------------------------

        if self.ir:
            self.ws.write(self.rowsize, 3, self.ir, self.style12)
        else:
            self.ws.write(self.rowsize, 3, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            self.ws.write(self.rowsize, 4, self.feedback_message, self.style12)
        else:
            self.ws.write(self.rowsize, 4, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.feedback and self.feedback['decisionText']:
            self.ws.write(self.rowsize, 5, self.feedback['decisionText'], self.style12)
        else:
            self.ws.write(self.rowsize, 5, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.feedback and self.feedback['partialFeedback'] == 1:
            self.ws.write(self.rowsize, 6, 'True', self.style12)
        elif self.feedback and self.feedback['partialFeedback'] == 0:
            self.ws.write(self.rowsize, 6, 'False', self.style12)
        # -------------------------------------------------------------------------------------------------------------

        if self.partial_data.get('message'):
            self.ws.write(self.rowsize, 7, self.partial_data['message'], self.style12)
        else:
            self.ws.write(self.rowsize, 7, None)
        # --------------------------------------------------------------------------------------------------------------
        if self.decision_update and self.decision_update.get('message'):
            self.ws.write(self.rowsize, 8, self.decision_update.get('message'), self.style12)
        elif self.decision_error and self.decision_error.get('errorDescription'):
            self.ws.write(self.rowsize, 8, self.decision_error.get('errorDescription'), self.style3)
        else:
            self.ws.write(self.rowsize, 8, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.decision_updated_feedback and self.decision_updated_feedback['decisionText']:
            self.ws.write(self.rowsize, 9, self.decision_updated_feedback['decisionText'], self.style12)
        else:
            self.ws.write(self.rowsize, 9, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_Datetime[loop] == self.feedback_data.get('interviewTime'):
            if self.xl_Datetime[loop] is None:
                self.ws.write(self.rowsize, 10, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 10, self.feedback_data.get('interviewTime'), self.style14)
        else:
            self.ws.write(self.rowsize, 10, self.feedback_data.get('interviewTime'), self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_int_datetime[loop] == self.feedback.get('interviewedTime'):
                if self.xl_int_datetime[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 11, self.feedback.get('interviewedTime'))
            else:
                self.ws.write(self.rowsize, 11, self.feedback.get('interviewedTime'), self.style14)
        elif self.ir:
            if self.xl_int_datetime[loop] is None:
                self.ws.write(self.rowsize, 11, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 11, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_id_01[loop] == self.skill_dict_1.get('skillId'):
                if self.xl_Skill_id_01[loop] is None:
                    self.ws.write(self.rowsize, 12, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 12, self.skill_dict_1.get('skillId'), self.style14)
        elif self.ir:
            if self.xl_Skill_id_01[loop] is None:
                self.ws.write(self.rowsize, 12, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 12, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_score_01[loop] == self.skill_dict_1.get('skillScore'):
                if self.xl_Skill_score_01[loop] is None:
                    self.ws.write(self.rowsize, 13, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 13, self.skill_dict_1.get('skillScore'), self.style14)
        elif self.ir:
            if self.xl_Skill_score_01[loop] is None:
                self.ws.write(self.rowsize, 13, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 13, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_id_02[loop] == self.skill_dict_2.get('skillId'):
                if self.xl_Skill_id_02[loop] is None:
                    self.ws.write(self.rowsize, 14, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 14, self.skill_dict_2.get('skillId'), self.style14)
        elif self.ir:
            if self.xl_Skill_id_02[loop] is None:
                self.ws.write(self.rowsize, 14, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 14, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_score_02[loop] == self.skill_dict_2.get('skillScore'):
                if self.xl_Skill_score_02[loop] is None:
                    self.ws.write(self.rowsize, 15, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 15, self.skill_dict_2.get('skillScore'), self.style14)
        elif self.ir:
            if self.xl_Skill_score_02[loop] is None:
                self.ws.write(self.rowsize, 15, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 15, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_id_03[loop] == self.skill_dict_3.get('skillId'):
                if self.xl_Skill_id_03[loop] is None:
                    self.ws.write(self.rowsize, 16, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 16, self.skill_dict_3.get('skillId'), self.style14)
        elif self.ir:
            if self.xl_Skill_id_03[loop] is None:
                self.ws.write(self.rowsize, 16, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 16, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_score_03[loop] == self.skill_dict_3.get('skillScore'):
                if self.xl_Skill_score_03[loop] is None:
                    self.ws.write(self.rowsize, 17, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 17, self.skill_dict_3.get('skillScore'), self.style14)
        elif self.ir:
            if self.xl_Skill_score_03[loop] is None:
                self.ws.write(self.rowsize, 17, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 17, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_id_04[loop] == self.skill_dict_4.get('skillId'):
                if self.xl_Skill_id_04[loop] is None:
                    self.ws.write(self.rowsize, 18, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 18, self.skill_dict_4.get('skillId'), self.style14)
        elif self.ir:
            if self.xl_Skill_id_04[loop] is None:
                self.ws.write(self.rowsize, 18, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 18, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Skill_score_04[loop] == self.skill_dict_4.get('skillScore'):
                if self.xl_Skill_score_04[loop] is None:
                    self.ws.write(self.rowsize, 19, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 19, self.skill_dict_4.get('skillScore'), self.style14)
        elif self.ir:
            if self.xl_Skill_score_04[loop] is None:
                self.ws.write(self.rowsize, 19, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 19, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_duration[loop] == self.feedback.get('duration'):
                if self.xl_duration[loop] is None:
                    self.ws.write(self.rowsize, 20, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 20, self.feedback.get('duration'), self.style14)
        elif self.ir:
            if self.xl_duration[loop] is None:
                self.ws.write(self.rowsize, 20, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 20, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_skill_comment[loop] == self.skill_dict_4.get('skillComment'):
                if self.xl_Skill_score_04[loop] is None:
                    self.ws.write(self.rowsize, 21, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 21, self.skill_dict_4.get('skillComment'), self.style14)
        elif self.ir:
            if self.xl_Skill_score_04[loop] is None:
                self.ws.write(self.rowsize, 21, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 21, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.is_feedback:
            if self.xl_Over_all_comment[loop] == self.feedback.get('comment'):
                if self.xl_Over_all_comment[loop] is None:
                    self.ws.write(self.rowsize, 22, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 22, self.feedback.get('comment'))
        elif self.ir:
            if self.xl_Over_all_comment[loop] is None:
                self.ws.write(self.rowsize, 22, 'Empty', self.style14)
        else:
            self.ws.write(self.rowsize, 22, None)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_update_skill_score_01[loop] == self.updated_skill_dict_1.get('skillScore'):
                if self.xl_update_skill_score_01[loop] is None:
                    self.ws.write(self.rowsize, 23, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 23, self.updated_skill_dict_1.get('skillScore'), self.style14)
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 23, self.updated_skill_dict_1.get('skillScore', 'No_Update_Details'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 23, self.updated_skill_dict_1.get('skillScore', 'NA'), self.style3)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 23,
                                  self.updated_skill_dict_1.get('skillScore', 'Not a Partial/Update feedback'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 23, self.updated_skill_dict_1.get('skillScore', 'NA'), self.style3)
        elif self.ir:
            if self.xl_update_skill_score_01[loop] is None:
                self.ws.write(self.rowsize, 23, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_updated_duration[loop] == self.updated_feedback.get('duration'):
                if self.xl_updated_duration[loop] is None:
                    self.ws.write(self.rowsize, 24, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 24, self.updated_feedback.get('duration'), self.style14)
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 24, self.updated_feedback.get('duration', 'No_Update_Details'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 24, self.updated_feedback.get('duration', 'NA'), self.style3)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 24,
                                  self.updated_feedback.get('duration', 'Not a Partial/Update feedback'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 24, self.updated_feedback.get('duration', 'NA'), self.style3)
        elif self.ir:
            if self.xl_updated_duration[loop] is None:
                self.ws.write(self.rowsize, 24, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_update_Skill_comment[loop] == self.updated_skill_dict_4.get('skillComment'):
                if self.xl_update_Skill_comment[loop] is None:
                    self.ws.write(self.rowsize, 25, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 25, self.updated_skill_dict_4.get('skillComment'), self.style14)
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 25, self.updated_skill_dict_4.get('skillComment', 'No_Update_Details'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 25, self.updated_skill_dict_4.get('skillComment', 'NA'), self.style3)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 25,
                                  self.updated_skill_dict_4.get('skillComment', 'Not a Partial/Update feedback'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 25, self.updated_skill_dict_4.get('skillComment', 'NA'), self.style3)
        elif self.ir:
            if self.xl_update_Skill_comment[loop] is None:
                self.ws.write(self.rowsize, 25, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_partial_feedback[loop] == 1 or self.xl_partial_feedback[loop] == 0:
            if self.xl_Updated_Over_all_comment[loop] == self.updated_feedback.get('comment'):
                if self.xl_Updated_Over_all_comment[loop] is None:
                    self.ws.write(self.rowsize, 26, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 26, self.updated_feedback.get('comment'), self.style14)
            elif self.xl_partial_feedback[loop] == 1:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 26, self.updated_feedback.get('comment', 'No_Update_Details'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 26, self.updated_feedback.get('comment', 'NA'), self.style3)
            elif self.xl_partial_feedback[loop] == 0:
                if self.feedback_message:
                    self.ws.write(self.rowsize, 26,
                                  self.updated_feedback.get('comment', 'Not a Partial/Update feedback'),
                                  self.style6)
                else:
                    self.ws.write(self.rowsize, 26, self.updated_feedback.get('comment', 'NA'), self.style3)
        elif self.ir:
            if self.xl_Updated_Over_all_comment[loop] is None:
                self.ws.write(self.rowsize, 26, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------

        if self.xl_exception_message[loop] is not None:
            if self.xl_exception_message[loop] == self.message:
                self.ws.write(self.rowsize, 27, self.message, self.style14)
            else:
                self.ws.write(self.rowsize, 27, self.message, self.style3)
        # --------------------------------------------------------------------------------------------------------------

        self.rowsize += 1  # Row increment
        Object.wb_Result.save(output_paths.outputpaths['Interview_flow_Output_sheet'])

        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)
        if self.success_case_03 == 'Pass':
            self.Actual_Success_case.append(self.success_case_03)
        if self.success_case_04 == 'Pass':
            self.Actual_Success_case.append(self.success_case_04)
        if self.success_case_05 == 'Pass':
            self.Actual_Success_case.append(self.success_case_05)
        if self.success_case_06 == 'Pass':
            self.Actual_Success_case.append(self.success_case_06)

    def overall_status(self):
        self.ws.write(0, 0, 'Provide Feedback', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Interview_flow_Output_sheet'])


Object = InterviewFeedback()
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
                Object.provide_feedback(looping)
                Object.feedback_details(looping)
                if Object.pf:
                    Object.partial_feedback(looping)
                    Object.updated_feedback_details(looping)
                if Object.updatedecision:
                    Object.update_decision(looping)
                    Object.decision_updated_feedback_details()
            Object.output_excel(looping)
            print('Excel data is ready')

            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.is_success = {}
            Object.message = {}
            Object.ir = {}
            Object.feedback_data = {}
            Object.feedback = {}
            Object.is_feedback = {}

            Object.updated_feedback_data = {}
            Object.updated_feedback = {}
            Object.partial_data = {}
            Object.feedback_message = {}
            Object.applicant_details = {}

            # ----------
            # Skill dict
            # ----------
            Object.skill_dict_1 = {}
            Object.skill_dict_2 = {}
            Object.skill_dict_3 = {}
            Object.skill_dict_4 = {}

            Object.filledFeedbackDetails = {}
            Object.skillAssessed_details = {}

            # ------------------
            # updated Skill dict
            # ------------------
            Object.updated_skill_dict_1 = {}
            Object.updated_skill_dict_2 = {}
            Object.updated_skill_dict_3 = {}
            Object.updated_skill_dict_4 = {}

            Object.updated_filledFeedbackDetails = {}
            Object.updated_skillAssessed_details = {}

            Object.decision_update = {}
            Object.updatedecision = {}
            Object.decision_error = {}
            Object.decision_updated_feedback = {}
            # ---------------------
            # Partial/updated  dict
            # ---------------------
            Object.pf = {}

            Object.success_case_01 = {}
            Object.success_case_02 = {}
            Object.success_case_03 = {}
            Object.success_case_04 = {}
            Object.success_case_05 = {}
            Object.success_case_06 = {}
    Object.overall_status()

except AttributeError as Object_error:
    print(Object_error)
