import json
import requests
import datetime
import xlrd
from hpro_automation import (api, login, input_paths, output_paths, work_book)


class UploadCandidate(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):
        self.start_time = str(datetime.datetime.now())
        super(UploadCandidate, self).__init__()

        # -------------------------------------------
        # Excel sheet headers for Output result sheet
        # -------------------------------------------

        header_column = 0
        excelheaders = ['Comparison', 'Actual_status', 'Candidate Id', 'Event Id', 'Event Name',
                        'Job Id', 'Job Name', 'Applicant Id', 'Test Id', 'Test Name',
                        'Original CId', 'Name', 'FirstName', 'MiddleName', 'LastName', 'Mobile1',
                        'PhoneOffice', 'Email1', 'Email2', 'Gender', 'MaritalStatus', 'DateOfBirth', 'USN', 'Address1',
                        'Address2', 'Final%', 'FinalEndYear', 'FinalDegree', 'FinalCollege', 'FinalDegreeType', '10th%',
                        '10thEndYear', '12th%', '12thEndYear', 'PanNo', 'PassportNo', 'CurrentLocation',
                        'TotalExperienceInMonths', 'Country', 'HierarchyId', 'Nationality', 'Sensitivity', 'StatusId',
                        'SourceId', 'CampusId', 'SourceType', 'Experience', 'EmployerId', 'DesignationId', 'Expertise',
                        'Notice Period', 'Integer1', 'Integer2', 'Integer3', 'Integer4', 'Integer5', 'Integer6',
                        'Integer7', 'Integer8', 'Integer9', 'Integer10', 'Integer11', 'Integer12', 'Integer13',
                        'Integer14', 'Integer15', 'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7',
                        'Text8', 'Text9', 'Text10', 'Text11', 'Text12', 'Text13', 'Text14', 'Text15', 'TextArea1',
                        'TextArea2', 'TextArea3', 'TextArea4', 'Expected_message']
        for headers in excelheaders:
            if headers in ['Comparison', 'Candidate Id', 'Original CId', 'Event Id', 'Event Name', 'Job Id',
                           'Job Name', 'Applicant Id', 'Overall_status', 'Message', 'Test Id', 'Test Name',
                           'Actual_status']:
                self.ws.write(1, header_column, headers, self.style2)
            else:
                self.ws.write(1, header_column, headers, self.style0)
            header_column += 1

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_eventId = []  # [] Initialising data from excel sheet to the variables
        self.xl_jobRoleId = []
        self.xl_mjrId = []
        self.xl_testId = []
        self.xl_Name = []
        self.xl_FirstName = []
        self.xl_MiddleName = []
        self.xl_LastName = []
        self.xl_Mobile1 = []
        self.xl_PhoneOffice = []
        self.xl_Email1 = []
        self.xl_Email2 = []
        self.xl_Gender = []
        self.xl_MaritalStatus = []
        self.xl_DateOfBirth = []
        self.xl_USN = []
        self.xl_Address1 = []
        self.xl_Address2 = []
        self.xl_PanNo = []
        self.xl_PassportNo = []
        self.xl_CurrentLocationId = []
        self.xl_TotalExperienceInMonths = []
        self.xl_Country = []
        self.xl_HierarchyId = []
        self.xl_Nationality = []
        self.xl_Sensitivity = []
        self.xl_StatusId = []
        self.xl_FinalPercentage = []
        self.xl_FinalEndYear = []
        self.xl_FinalDegreeId = []
        self.xl_FinalCollegeId = []
        self.xl_FinalDegreeTypeId = []
        self.xl_10thDegreeId = []
        self.xl_10thPercentage = []
        self.xl_10thEndYear = []
        self.xl_12thDegreeId = []
        self.xl_12thPercentage = []
        self.xl_12thEndYear = []
        self.xl_SourceId = []
        self.xl_CampusId = []
        self.xl_SourceType = []
        self.xl_Experience = []
        self.xl_EmployerId = []
        self.xl_DesignationId = []
        self.xl_Expertise = []
        self.xl_NoticePeriod = []
        self.xl_Integer1 = []
        self.xl_Integer2 = []
        self.xl_Integer3 = []
        self.xl_Integer4 = []
        self.xl_Integer5 = []
        self.xl_Integer6 = []
        self.xl_Integer7 = []
        self.xl_Integer8 = []
        self.xl_Integer9 = []
        self.xl_Integer10 = []
        self.xl_Integer11 = []
        self.xl_Integer12 = []
        self.xl_Integer13 = []
        self.xl_Integer14 = []
        self.xl_Integer15 = []
        self.xl_Text1 = []
        self.xl_Text2 = []
        self.xl_Text3 = []
        self.xl_Text4 = []
        self.xl_Text5 = []
        self.xl_Text6 = []
        self.xl_Text7 = []
        self.xl_Text8 = []
        self.xl_Text9 = []
        self.xl_Text10 = []
        self.xl_Text11 = []
        self.xl_Text12 = []
        self.xl_Text13 = []
        self.xl_Text14 = []
        self.xl_Text15 = []
        self.xl_TextArea1 = []
        self.xl_TextArea2 = []
        self.xl_TextArea3 = []
        self.xl_TextArea4 = []
        self.xl_Exception_message = []

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 66)))
        self.Actual_Success_case = []

        # -------------------------------------------------------------------------------------------------------
        # Dictionary for candidate_get_by_id_details, candidate_educational_details, candidate_experience_details
        # --------------------------------------------------------------------------------------------------------
        self.personal_details_dict = {}
        self.candidate_personal_details = self.personal_details_dict
        self.source_details_dict = {}
        self.candidate_source_details = self.source_details_dict
        self.custom_details_dict = {}
        self.candidate_custom_details = self.custom_details_dict
        self.final_degree_dict = {}
        self.candidate_final_degree_dict = self.final_degree_dict
        self.tenth_dict = {}
        self.candidate_tenth_dict = self.tenth_dict
        self.twelfth_dict = {}
        self.candidate_twelfth_dict = self.twelfth_dict
        self.experience_dict = {}
        self.candidate_experience_dict = self.experience_dict
        self.app_dict = {}
        self.event_applicant_dict = self.app_dict
        self.test_dict = {}
        self.test_detail = self.test_dict
        self.isCreated = {}
        self.isc = self.isCreated
        self.OrginalCID = {}
        self.O_cid = self.OrginalCID
        self.message = {}
        self.me = self.message
        self.CID = {}
        self.c_id = self.CID
        self.candidatesavemessage = {}
        self.c_s_m = self.candidatesavemessage
        self.applicantDetails = {}
        self.app_details = self.applicantDetails
        self.partialmessage = {}
        self.p_m = self.partialmessage

        self.success_case_01 = {}
        self.success_case_02 = {}
        self.success_case_03 = {}
        self.success_case_04 = {}
        self.success_case_05 = {}
        self.failure_case_01 = {}
        self.failure_case_02 = {}
        self.headers = {}

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook(input_paths.inputpaths['UploadCandi_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            # ------------------------------
            # Event, Job, Mjr, Test details
            # ------------------------------
            if not rows[0]:
                self.xl_eventId.append(None)
            else:
                self.xl_eventId.append(int(rows[0]))

            if not rows[1]:
                self.xl_jobRoleId.append(None)
            else:
                self.xl_jobRoleId.append(int(rows[1]))

            if not rows[2]:
                self.xl_mjrId.append(None)
            else:
                self.xl_mjrId.append(int(rows[2]))

            if not rows[3]:
                self.xl_testId.append(None)
            else:
                self.xl_testId.append(int(rows[3]))

            #  ----------------------------------------------------------
            # Personal, Source, Educational, Experience, Custom details
            # ----------------------------------------------------------
            if not rows[4]:
                self.xl_Name.append(None)
            else:
                self.xl_Name.append(rows[4])

            if not rows[5]:
                self.xl_FirstName.append(None)
            else:
                self.xl_FirstName.append(rows[5])

            if not rows[6]:
                self.xl_MiddleName.append(None)
            else:
                self.xl_MiddleName.append(rows[6])

            if not rows[7]:
                self.xl_LastName.append(None)
            else:
                self.xl_LastName.append(rows[7])

            if not rows[8]:
                self.xl_Mobile1.append(None)
            else:
                self.xl_Mobile1.append(str(int(rows[8])))

            if not rows[9]:
                self.xl_PhoneOffice.append(None)
            else:
                self.xl_PhoneOffice.append(str(int(rows[9])))

            if not rows[10]:
                self.xl_Email1.append(None)
            else:
                self.xl_Email1.append(str(rows[10]))

            if not rows[11]:
                self.xl_Email2.append(None)
            else:
                self.xl_Email2.append(str(rows[11]))

            if not rows[12]:
                self.xl_Gender.append(None)
            else:
                self.xl_Gender.append(int(rows[12]))

            if not rows[13]:
                self.xl_MaritalStatus.append(None)
            else:
                self.xl_MaritalStatus.append(int(rows[13]))

            if not rows[14]:
                self.xl_DateOfBirth.append(None)
            else:
                self.xl_DateOfBirth.append(str(rows[14]))

            if not rows[15]:
                self.xl_USN.append(None)
            else:
                self.xl_USN.append(str(rows[15]))

            if not rows[16]:
                self.xl_Address1.append(None)
            else:
                self.xl_Address1.append(str(rows[16]))

            if not rows[17]:
                self.xl_Address2.append(None)
            else:
                self.xl_Address2.append(str(rows[17]))

            if not rows[18]:
                self.xl_PanNo.append(None)
            else:
                self.xl_PanNo.append(str(rows[18]))

            if not rows[19]:
                self.xl_PassportNo.append(None)
            else:
                self.xl_PassportNo.append(str(rows[19]))

            if not rows[20]:
                self.xl_CurrentLocationId.append(None)
            else:
                self.xl_CurrentLocationId.append(int(rows[20]))

            if not rows[21]:
                self.xl_TotalExperienceInMonths.append(None)
            else:
                self.xl_TotalExperienceInMonths.append(int(rows[21]))

            if not rows[22]:
                self.xl_Country.append(None)
            else:
                self.xl_Country.append(int(rows[22]))

            if not rows[23]:
                self.xl_HierarchyId.append(None)
            else:
                self.xl_HierarchyId.append(int(rows[23]))

            if not rows[24]:
                self.xl_Nationality.append(None)
            else:
                self.xl_Nationality.append(int(rows[24]))

            if not rows[25]:
                self.xl_Sensitivity.append(None)
            else:
                self.xl_Sensitivity.append(int(rows[25]))

            if not rows[26]:
                self.xl_StatusId.append(None)
            else:
                self.xl_StatusId.append(int(rows[26]))

            if not rows[27]:
                self.xl_FinalPercentage.append(None)
            else:
                self.xl_FinalPercentage.append(float(rows[27]))

            if not rows[28]:
                self.xl_FinalEndYear.append(None)
            else:
                self.xl_FinalEndYear.append(int(rows[28]))

            if not rows[29]:
                self.xl_FinalDegreeId.append(None)
            else:
                self.xl_FinalDegreeId.append(int(rows[29]))

            if not rows[30]:
                self.xl_FinalCollegeId.append(None)
            else:
                self.xl_FinalCollegeId.append(int(rows[30]))

            if not rows[31]:
                self.xl_FinalDegreeTypeId.append(None)
            else:
                self.xl_FinalDegreeTypeId.append(int(rows[31]))

            if not rows[32]:
                self.xl_10thDegreeId.append(None)
            else:
                self.xl_10thDegreeId.append(int(rows[32]))

            if not rows[33]:
                self.xl_10thPercentage.append(None)
            else:
                self.xl_10thPercentage.append(float(rows[33]))

            if not rows[34]:
                self.xl_10thEndYear.append(None)
            else:
                self.xl_10thEndYear.append(int(rows[34]))

            if not rows[35]:
                self.xl_12thDegreeId.append(None)
            else:
                self.xl_12thDegreeId.append(int(rows[35]))

            if not rows[36]:
                self.xl_12thPercentage.append(None)
            else:
                self.xl_12thPercentage.append(float(rows[36]))

            if not rows[37]:
                self.xl_12thEndYear.append(None)
            else:
                self.xl_12thEndYear.append(int(rows[37]))

            if not rows[38]:
                self.xl_SourceId.append(None)
            else:
                self.xl_SourceId.append(int(rows[38]))

            if not rows[39]:
                self.xl_CampusId.append(None)
            else:
                self.xl_CampusId.append(int(rows[39]))

            if not rows[40]:
                self.xl_SourceType.append(None)
            else:
                self.xl_SourceType.append(int(rows[40]))

            if not rows[41]:
                self.xl_Experience.append(None)
            else:
                self.xl_Experience.append(int(rows[41]))

            if not rows[42]:
                self.xl_EmployerId.append(None)
            else:
                self.xl_EmployerId.append(int(rows[42]))

            if not rows[43]:
                self.xl_DesignationId.append(None)
            else:
                self.xl_DesignationId.append(int(rows[43]))

            if not rows[44]:
                self.xl_Expertise.append(None)
            else:
                self.xl_Expertise.append(int(rows[44]))

            if not rows[45]:
                self.xl_NoticePeriod.append(None)
            else:
                self.xl_NoticePeriod.append(int(rows[45]))

            if not rows[46]:
                self.xl_Integer1.append(None)
            else:
                self.xl_Integer1.append(int(rows[46]))

            if not rows[47]:
                self.xl_Integer2.append(None)
            else:
                self.xl_Integer2.append(int(rows[47]))

            if not rows[48]:
                self.xl_Integer3.append(None)
            else:
                self.xl_Integer3.append(int(rows[48]))

            if not rows[49]:
                self.xl_Integer4.append(None)
            else:
                self.xl_Integer4.append(int(rows[49]))

            if not rows[50]:
                self.xl_Integer5.append(None)
            else:
                self.xl_Integer5.append(int(rows[50]))

            if not rows[51]:
                self.xl_Integer6.append(None)
            else:
                self.xl_Integer6.append(int(rows[51]))

            if not rows[52]:
                self.xl_Integer7.append(None)
            else:
                self.xl_Integer7.append(int(rows[52]))

            if not rows[53]:
                self.xl_Integer8.append(None)
            else:
                self.xl_Integer8.append(int(rows[53]))

            if not rows[54]:
                self.xl_Integer9.append(None)
            else:
                self.xl_Integer9.append(int(rows[54]))

            if not rows[55]:
                self.xl_Integer10.append(None)
            else:
                self.xl_Integer10.append(int(rows[55]))

            if not rows[56]:
                self.xl_Integer11.append(None)
            else:
                self.xl_Integer11.append(int(rows[56]))

            if not rows[57]:
                self.xl_Integer12.append(None)
            else:
                self.xl_Integer12.append(int(rows[57]))

            if not rows[58]:
                self.xl_Integer13.append(None)
            else:
                self.xl_Integer13.append(int(rows[58]))

            if not rows[59]:
                self.xl_Integer14.append(None)
            else:
                self.xl_Integer14.append(int(rows[59]))

            if not rows[60]:
                self.xl_Integer15.append(None)
            else:
                self.xl_Integer15.append(int(rows[60]))

            if not rows[61]:
                self.xl_Text1.append(None)
            else:
                self.xl_Text1.append(rows[61])

            if not rows[62]:
                self.xl_Text2.append(None)
            else:
                self.xl_Text2.append(rows[62])

            if not rows[63]:
                self.xl_Text3.append(None)
            else:
                self.xl_Text3.append(rows[63])

            if not rows[64]:
                self.xl_Text4.append(None)
            else:
                self.xl_Text4.append(rows[64])

            if not rows[65]:
                self.xl_Text5.append(None)
            else:
                self.xl_Text5.append(rows[65])

            if not rows[66]:
                self.xl_Text6.append(None)
            else:
                self.xl_Text6.append(rows[66])

            if not rows[67]:
                self.xl_Text7.append(None)
            else:
                self.xl_Text7.append(rows[67])

            if not rows[68]:
                self.xl_Text8.append(None)
            else:
                self.xl_Text8.append(rows[68])

            if not rows[69]:
                self.xl_Text9.append(None)
            else:
                self.xl_Text9.append(rows[69])

            if not rows[70]:
                self.xl_Text10.append(None)
            else:
                self.xl_Text10.append(rows[70])

            if not rows[71]:
                self.xl_Text11.append(None)
            else:
                self.xl_Text11.append(rows[71])

            if not rows[72]:
                self.xl_Text12.append(None)
            else:
                self.xl_Text12.append(rows[72])

            if not rows[73]:
                self.xl_Text13.append(None)
            else:
                self.xl_Text13.append(rows[73])

            if not rows[74]:
                self.xl_Text14.append(None)
            else:
                self.xl_Text14.append(rows[74])

            if not rows[75]:
                self.xl_Text15.append(None)
            else:
                self.xl_Text15.append(rows[75])

            if not rows[76]:
                self.xl_TextArea1.append(None)
            else:
                self.xl_TextArea1.append(rows[76])

            if not rows[77]:
                self.xl_TextArea2.append(None)
            else:
                self.xl_TextArea2.append(rows[77])

            if not rows[78]:
                self.xl_TextArea3.append(None)
            else:
                self.xl_TextArea3.append(rows[78])

            if not rows[79]:
                self.xl_TextArea4.append(None)
            else:
                self.xl_TextArea4.append(rows[79])

            if not rows[80]:
                self.xl_Exception_message.append(None)
            else:
                self.xl_Exception_message.append(rows[80])

    def bulk_create_tag_candidates(self, iteration):

        self.lambda_function('bulkCreateTagCandidates')
        self.headers['APP-NAME'] = 'crpo'

        # -------------------------
        # Candidate create request
        # -------------------------
        create_candidate_request = {"createTagCandidates": [{
            "PersonalDetails": {
                "Name": self.xl_Name[iteration],
                "FirstName": self.xl_FirstName[iteration],
                "MiddleName": self.xl_MiddleName[iteration],
                "LastName": self.xl_LastName[iteration],
                "Mobile1": self.xl_Mobile1[iteration],
                "PhoneOffice": self.xl_PhoneOffice[iteration],
                "Email1": self.xl_Email1[iteration],
                "Email2": self.xl_Email2[iteration],
                "Gender": self.xl_Gender[iteration],
                "DateOfBirth": self.xl_DateOfBirth[iteration],
                "USN": self.xl_USN[iteration],
                "MaritalStatus": self.xl_MaritalStatus[iteration],
                "Address1": self.xl_Address1[iteration],
                "Address2": self.xl_Address2[iteration],
                "PanNo": self.xl_PanNo[iteration],
                "PassportNo": self.xl_PassportNo[iteration],
                "CurrentLocationId": self.xl_CurrentLocationId[iteration],
                "TotalExperienceInMonths": self.xl_TotalExperienceInMonths[iteration],
                "Country": self.xl_Country[iteration],
                "HierarchyId": self.xl_HierarchyId[iteration],
                "Nationality": self.xl_Nationality[iteration],
                "Sensitivity": self.xl_Sensitivity[iteration],
                "StatusId": self.xl_StatusId[iteration],
                "ExpertiseId1": self.xl_Expertise[iteration]
            },
            "EducationDetails": {
                "AddedItems": [{
                    "IsPercentage": True,
                    "Percentage": self.xl_FinalPercentage[iteration],
                    "EndYear": self.xl_FinalEndYear[iteration],
                    "IsFinal": True,
                    "DegreeId": self.xl_FinalDegreeId[iteration],
                    "CollegeId": self.xl_FinalCollegeId[iteration],
                    "DegreeTypeId": self.xl_FinalDegreeTypeId[iteration]
                }, {
                    "IsPercentage": True,
                    "DegreeId": self.xl_10thDegreeId[iteration],
                    "Percentage": self.xl_10thPercentage[iteration],
                    "EndYear": self.xl_10thEndYear[iteration],
                    "IsFinal": False
                }, {
                    "IsPercentage": False,
                    "DegreeId": self.xl_12thDegreeId[iteration],
                    "Percentage": self.xl_12thPercentage[iteration],
                    "EndYear": self.xl_12thEndYear[iteration],
                    "IsFinal": False
                }]},
            "ExperienceDetails": {
                "AddedItems": [{
                    "IsLatest": True,
                    "Experience": self.xl_Experience[iteration],
                    "EmployerId": self.xl_EmployerId[iteration],
                    "DesignationId": self.xl_DesignationId[iteration]
                }]
            },
            "CustomDetails": {
                "Integer1": self.xl_Integer1[iteration],
                "Integer2": self.xl_Integer2[iteration],
                "Integer3": self.xl_Integer3[iteration],
                "Integer4": self.xl_Integer4[iteration],
                "Integer5": self.xl_Integer5[iteration],
                "Integer6": self.xl_Integer6[iteration],
                "Integer7": self.xl_Integer7[iteration],
                "Integer8": self.xl_Integer8[iteration],
                "Integer9": self.xl_Integer9[iteration],
                "Integer10": self.xl_Integer10[iteration],
                "Integer11": self.xl_Integer11[iteration],
                "Integer12": self.xl_Integer12[iteration],
                "Integer13": self.xl_Integer13[iteration],
                "Integer14": self.xl_Integer14[iteration],
                "Integer15": self.xl_Integer15[iteration],
                "Text1": self.xl_Text1[iteration],
                "Text2": self.xl_Text2[iteration],
                "Text3": self.xl_Text3[iteration],
                "Text4": self.xl_Text4[iteration],
                "Text5": self.xl_Text5[iteration],
                "Text6": self.xl_Text6[iteration],
                "Text7": self.xl_Text7[iteration],
                "Text8": self.xl_Text8[iteration],
                "Text9": self.xl_Text9[iteration],
                "Text10": self.xl_Text10[iteration],
                "Text11": self.xl_Text11[iteration],
                "Text12": self.xl_Text12[iteration],
                "Text13": self.xl_Text13[iteration],
                "Text14": self.xl_Text14[iteration],
                "Text15": self.xl_Text15[iteration],
                "TextArea1": self.xl_TextArea1[iteration],
                "TextArea2": self.xl_TextArea2[iteration],
                "TextArea3": self.xl_TextArea3[iteration],
                "TextArea4": self.xl_TextArea4[iteration]
            },
            "PreferenceDetails": {
                "NoticePeriod": self.xl_NoticePeriod[iteration]
            },
            "SourceDetails": {
                "SourceId": self.xl_SourceId[iteration],
                "CampusId": self.xl_CampusId[iteration],
                "SourceType": self.xl_SourceType[iteration]
            },
            "applicantDetail": {
                "eventId": self.xl_eventId[iteration],
                "jobRoleId": self.xl_jobRoleId[iteration],
                "mjrId": self.xl_mjrId[iteration],
                "testId": [self.xl_testId[iteration]],
                "isCreateDuplicate": False
            }
        }],
            "Sync": "True"
        }
        create_candidate = requests.post(api.web_api['bulkCreateTagCandidates'], headers=self.headers,
                                         data=json.dumps(create_candidate_request, default=str), verify=False)
        print(create_candidate.headers)
        create_candidate_response_dict = json.loads(create_candidate.content)
        candidate_response_data = create_candidate_response_dict['data']

        # -----------------------------------------
        # API response from bulkCreateTagCandidate
        # -----------------------------------------
        for response in candidate_response_data:
            self.isCreated = response['isCreated']
            self.OrginalCID = response.get('originalCandidateId')
            self.message = response.get('duplicateCandidateMessage')
            self.CID = response.get('candidateId')
            self.candidatesavemessage = response.get('candidateSaveMessage')
            self.applicantDetails = response.get('applicantDetails')
            successmessage = self.applicantDetails.get('success') if self.applicantDetails else None
            listmessage = successmessage.get(str(self.CID)) if successmessage else None

            if listmessage == listmessage:
                if listmessage is None:
                    self.partialmessage = None
                else:
                    self.partialmessage = ', '.join(listmessage)
            print(self.partialmessage)

            if self.isCreated:  # Always Boolean is true
                print("Create Candidate :", self.isCreated)
                print("candidate Id ::", self.CID)
            else:
                print("Create Candidate ::", self.isCreated)
                print("Message ::", self.message)

    def candidate_get_by_id_details(self):

        self.lambda_function('CandidateGetbyId')
        self.headers['APP-NAME'] = 'crpo'

        get_candidate_details = requests.post(api.web_api['CandidateGetbyId'].format(self.CID), headers=self.headers)
        print(get_candidate_details.headers)
        candidate_details = json.loads(get_candidate_details.content)
        candidate_dict = candidate_details['Candidate']
        self.personal_details_dict = candidate_dict['PersonalDetails']
        self.source_details_dict = candidate_dict['SourceDetails']
        self.custom_details_dict = candidate_dict['CustomDetails']

    def candidate_educational_details(self, loop):

        self.lambda_function('Candidate_Educationaldetails')
        self.headers['APP-NAME'] = 'crpo'

        get_educational_details = requests.post(api.web_api['Candidate_Educationaldetails'].format(self.CID),
                                                headers=self.headers)
        print(get_educational_details.headers)
        educational_details = json.loads(get_educational_details.content)
        educational_dict = educational_details['EducationProfile']
        for edu in educational_dict:
            if edu['DegreeId'] == self.xl_FinalDegreeId[loop]:
                self.final_degree_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_FinalDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_10thDegreeId[loop]:
                self.tenth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_10thDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_12thDegreeId[loop]:
                self.twelfth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_12thDegreeId[loop]), None)

    def candidate_experience_details(self):

        self.lambda_function('Candidate_ExperienceDetails')
        self.headers['APP-NAME'] = 'crpo'

        get_experience_details = requests.post(api.web_api['Candidate_ExperienceDetails'].format(self.CID),
                                               headers=self.headers)
        print(get_experience_details.headers)
        experience_details = json.loads(get_experience_details.content)
        experience_dict = experience_details['WorkProfile']
        for exp in experience_dict:
            self.experience_dict = exp

    def event_applicants(self, loop):

        self.lambda_function('getAllApplicants')
        self.headers['APP-NAME'] = 'crpo'

        eventapplicant_request = {
            "RecruitEventId": self.xl_eventId[loop],
            "PagingCriteriaType": {
                "MaxResults": 1000,
                "PageNumber": 1
            }
        }
        eventapplicant_api = requests.post(api.web_api['getAllApplicants'], headers=self.headers,
                                           data=json.dumps(eventapplicant_request, default=str), verify=False)
        print(eventapplicant_api.headers)
        applicant_dict = json.loads(eventapplicant_api.content)
        print(applicant_dict)
        applicant_data = applicant_dict['data']
        if applicant_data:
            for appdata in applicant_data:
                # -----------------------------------
                # Matching with created candidate Id
                # -----------------------------------
                if appdata['CandidateId'] == self.CID:
                    self.app_dict = next((item for item in applicant_data if item['CandidateId'] == self.CID), None)
                    test_details = self.app_dict['TestUserDetailType']
                    print(test_details)
                    if test_details:
                        for td in test_details:
                            if td['TestId'] == self.xl_testId[loop]:
                                self.test_dict = next(
                                    (item for item in test_details if item['TestId'] == self.xl_testId[loop]), None)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)
        if self.xl_Name[loop]:
            self.ws.write(self.rowsize, 11, self.xl_Name[loop])
        else:
            self.ws.write(self.rowsize, 11, 'Empty')
        if self.xl_FirstName[loop]:
            self.ws.write(self.rowsize, 12, self.xl_FirstName[loop])
        else:
            self.ws.write(self.rowsize, 12, 'Empty')
        if self.xl_MiddleName[loop]:
            self.ws.write(self.rowsize, 13, self.xl_MiddleName[loop])
        else:
            self.ws.write(self.rowsize, 13, 'Empty')
        if self.xl_LastName[loop]:
            self.ws.write(self.rowsize, 14, self.xl_LastName[loop])
        else:
            self.ws.write(self.rowsize, 14, 'Empty')
        if self.xl_Mobile1[loop]:
            self.ws.write(self.rowsize, 15, self.xl_Mobile1[loop])
        else:
            self.ws.write(self.rowsize, 15, 'Empty')
        if self.xl_PhoneOffice[loop]:
            self.ws.write(self.rowsize, 16, self.xl_PhoneOffice[loop])
        else:
            self.ws.write(self.rowsize, 16, 'Empty')
        if self.xl_Email1[loop]:
            self.ws.write(self.rowsize, 17, self.xl_Email1[loop])
        else:
            self.ws.write(self.rowsize, 17, 'Empty')
        if self.xl_Email2[loop]:
            self.ws.write(self.rowsize, 18, self.xl_Email2[loop])
        else:
            self.ws.write(self.rowsize, 18, 'Empty')
        if self.xl_Gender[loop]:
            self.ws.write(self.rowsize, 19, self.xl_Gender[loop])
        else:
            self.ws.write(self.rowsize, 19, 'Empty')
        if self.xl_MaritalStatus[loop]:
            self.ws.write(self.rowsize, 20, self.xl_MaritalStatus[loop])
        else:
            self.ws.write(self.rowsize, 20, 'Empty')
        if self.xl_DateOfBirth[loop]:
            self.ws.write(self.rowsize, 21, self.xl_DateOfBirth[loop])
        else:
            self.ws.write(self.rowsize, 21, 'Empty')
        if self.xl_USN[loop]:
            self.ws.write(self.rowsize, 22, self.xl_USN[loop])
        else:
            self.ws.write(self.rowsize, 22, 'Empty')
        if self.xl_Address1[loop]:
            self.ws.write(self.rowsize, 23, self.xl_Address1[loop])
        else:
            self.ws.write(self.rowsize, 23, 'Empty')
        if self.xl_Address2[loop]:
            self.ws.write(self.rowsize, 24, self.xl_Address2[loop])
        else:
            self.ws.write(self.rowsize, 24, 'Empty')
        if self.xl_FinalPercentage[loop]:
            self.ws.write(self.rowsize, 25, self.xl_FinalPercentage[loop])
        else:
            self.ws.write(self.rowsize, 25, 'Empty')
        if self.xl_FinalEndYear[loop]:
            self.ws.write(self.rowsize, 26, self.xl_FinalEndYear[loop])
        else:
            self.ws.write(self.rowsize, 26, 'Empty')
        if self.xl_FinalDegreeId[loop]:
            self.ws.write(self.rowsize, 27, self.xl_FinalDegreeId[loop])
        else:
            self.ws.write(self.rowsize, 27, 'Empty')
        if self.xl_FinalCollegeId[loop]:
            self.ws.write(self.rowsize, 28, self.xl_FinalCollegeId[loop])
        else:
            self.ws.write(self.rowsize, 28, 'Empty')
        if self.xl_FinalDegreeTypeId[loop]:
            self.ws.write(self.rowsize, 29, self.xl_FinalDegreeTypeId[loop])
        else:
            self.ws.write(self.rowsize, 29, 'Empty')
        if self.xl_10thPercentage[loop]:
            self.ws.write(self.rowsize, 30, self.xl_10thPercentage[loop])
        else:
            self.ws.write(self.rowsize, 30, 'Empty')
        if self.xl_10thEndYear[loop]:
            self.ws.write(self.rowsize, 31, self.xl_10thEndYear[loop])
        else:
            self.ws.write(self.rowsize, 31, 'Empty')
        if self.xl_12thPercentage[loop]:
            self.ws.write(self.rowsize, 32, self.xl_12thPercentage[loop])
        else:
            self.ws.write(self.rowsize, 32, 'Empty')
        if self.xl_12thEndYear[loop]:
            self.ws.write(self.rowsize, 33, self.xl_12thEndYear[loop])
        else:
            self.ws.write(self.rowsize, 33, 'Empty')
        if self.xl_PanNo[loop]:
            self.ws.write(self.rowsize, 34, self.xl_PanNo[loop])
        else:
            self.ws.write(self.rowsize, 34, 'Empty')
        if self.xl_PassportNo[loop]:
            self.ws.write(self.rowsize, 35, self.xl_PassportNo[loop])
        else:
            self.ws.write(self.rowsize, 35, 'Empty')
        if self.xl_CurrentLocationId[loop]:
            self.ws.write(self.rowsize, 36, self.xl_CurrentLocationId[loop])
        else:
            self.ws.write(self.rowsize, 36, 'Empty')
        if self.xl_TotalExperienceInMonths[loop]:
            self.ws.write(self.rowsize, 37, self.xl_TotalExperienceInMonths[loop])
        else:
            self.ws.write(self.rowsize, 37, 'Empty')
        if self.xl_Country[loop]:
            self.ws.write(self.rowsize, 38, self.xl_Country[loop])
        else:
            self.ws.write(self.rowsize, 38, 'Empty')
        if self.xl_HierarchyId[loop]:
            self.ws.write(self.rowsize, 39, self.xl_HierarchyId[loop])
        else:
            self.ws.write(self.rowsize, 39, 'Empty')
        if self.xl_Nationality[loop]:
            self.ws.write(self.rowsize, 40, self.xl_Nationality[loop])
        else:
            self.ws.write(self.rowsize, 40, 'Empty')
        if self.xl_Sensitivity[loop]:
            self.ws.write(self.rowsize, 41, self.xl_Sensitivity[loop])
        else:
            self.ws.write(self.rowsize, 41, 'Empty')
        if self.xl_StatusId[loop]:
            self.ws.write(self.rowsize, 42, self.xl_StatusId[loop])
        else:
            self.ws.write(self.rowsize, 42, 'Empty')
        if self.xl_SourceId[loop]:
            self.ws.write(self.rowsize, 43, self.xl_SourceId[loop])
        else:
            self.ws.write(self.rowsize, 43, 'Empty')
        if self.xl_CampusId[loop]:
            self.ws.write(self.rowsize, 44, self.xl_CampusId[loop])
        else:
            self.ws.write(self.rowsize, 44, 'Empty')
        if self.xl_SourceType[loop]:
            self.ws.write(self.rowsize, 45, self.xl_SourceType[loop])
        else:
            self.ws.write(self.rowsize, 45, 'Empty')
        if self.xl_Experience[loop]:
            self.ws.write(self.rowsize, 46, self.xl_Experience[loop])
        else:
            self.ws.write(self.rowsize, 46, 'Empty')
        if self.xl_EmployerId[loop]:
            self.ws.write(self.rowsize, 47, self.xl_EmployerId[loop])
        else:
            self.ws.write(self.rowsize, 47, 'Empty')
        if self.xl_DesignationId[loop]:
            self.ws.write(self.rowsize, 48, self.xl_DesignationId[loop])
        else:
            self.ws.write(self.rowsize, 48, 'Empty')
        if self.xl_Expertise[loop]:
            self.ws.write(self.rowsize, 49, self.xl_Expertise[loop])
        else:
            self.ws.write(self.rowsize, 49, 'Empty')
        if self.xl_NoticePeriod[loop]:
            self.ws.write(self.rowsize, 50, self.xl_NoticePeriod[loop])
        else:
            self.ws.write(self.rowsize, 50, 'Empty')
        if self.xl_Integer1[loop]:
            self.ws.write(self.rowsize, 51, self.xl_Integer1[loop])
        else:
            self.ws.write(self.rowsize, 51, 'Empty')
        if self.xl_Integer2[loop]:
            self.ws.write(self.rowsize, 52, self.xl_Integer2[loop])
        else:
            self.ws.write(self.rowsize, 52, 'Empty')
        if self.xl_Integer3[loop]:
            self.ws.write(self.rowsize, 53, self.xl_Integer3[loop])
        else:
            self.ws.write(self.rowsize, 53, 'Empty')
        if self.xl_Integer4[loop]:
            self.ws.write(self.rowsize, 54, self.xl_Integer4[loop])
        else:
            self.ws.write(self.rowsize, 54, 'Empty')
        if self.xl_Integer5[loop]:
            self.ws.write(self.rowsize, 55, self.xl_Integer5[loop])
        else:
            self.ws.write(self.rowsize, 55, 'Empty')
        if self.xl_Integer6[loop]:
            self.ws.write(self.rowsize, 56, self.xl_Integer6[loop])
        else:
            self.ws.write(self.rowsize, 56, 'Empty')
        if self.xl_Integer7[loop]:
            self.ws.write(self.rowsize, 57, self.xl_Integer7[loop])
        else:
            self.ws.write(self.rowsize, 57, 'Empty')
        if self.xl_Integer8[loop]:
            self.ws.write(self.rowsize, 58, self.xl_Integer8[loop])
        else:
            self.ws.write(self.rowsize, 58, 'Empty')
        if self.xl_Integer9[loop]:
            self.ws.write(self.rowsize, 59, self.xl_Integer9[loop])
        else:
            self.ws.write(self.rowsize, 59, 'Empty')
        if self.xl_Integer10[loop]:
            self.ws.write(self.rowsize, 60, self.xl_Integer10[loop])
        else:
            self.ws.write(self.rowsize, 60, 'Empty')
        if self.xl_Integer11[loop]:
            self.ws.write(self.rowsize, 61, self.xl_Integer11[loop])
        else:
            self.ws.write(self.rowsize, 61, 'Empty')
        if self.xl_Integer12[loop]:
            self.ws.write(self.rowsize, 62, self.xl_Integer12[loop])
        else:
            self.ws.write(self.rowsize, 62, 'Empty')
        if self.xl_Integer13[loop]:
            self.ws.write(self.rowsize, 63, self.xl_Integer13[loop])
        else:
            self.ws.write(self.rowsize, 63, 'Empty')
        if self.xl_Integer14[loop]:
            self.ws.write(self.rowsize, 64, self.xl_Integer14[loop])
        else:
            self.ws.write(self.rowsize, 64, 'Empty')
        if self.xl_Integer15[loop]:
            self.ws.write(self.rowsize, 65, self.xl_Integer15[loop])
        else:
            self.ws.write(self.rowsize, 65, 'Empty')
        if self.xl_Text1[loop]:
            self.ws.write(self.rowsize, 66, self.xl_Text1[loop])
        else:
            self.ws.write(self.rowsize, 66, 'Empty')
        if self.xl_Text2[loop]:
            self.ws.write(self.rowsize, 67, self.xl_Text2[loop])
        else:
            self.ws.write(self.rowsize, 67, 'Empty')
        if self.xl_Text3[loop]:
            self.ws.write(self.rowsize, 68, self.xl_Text3[loop])
        else:
            self.ws.write(self.rowsize, 68, 'Empty')
        if self.xl_Text4[loop]:
            self.ws.write(self.rowsize, 69, self.xl_Text4[loop])
        else:
            self.ws.write(self.rowsize, 69, 'Empty')
        if self.xl_Text5[loop]:
            self.ws.write(self.rowsize, 70, self.xl_Text5[loop])
        else:
            self.ws.write(self.rowsize, 70, 'Empty')
        if self.xl_Text6[loop]:
            self.ws.write(self.rowsize, 71, self.xl_Text6[loop])
        else:
            self.ws.write(self.rowsize, 71, 'Empty')
        if self.xl_Text7[loop]:
            self.ws.write(self.rowsize, 72, self.xl_Text7[loop])
        else:
            self.ws.write(self.rowsize, 72, 'Empty')
        if self.xl_Text8[loop]:
            self.ws.write(self.rowsize, 73, self.xl_Text8[loop])
        else:
            self.ws.write(self.rowsize, 73, 'Empty')
        if self.xl_Text9[loop]:
            self.ws.write(self.rowsize, 74, self.xl_Text9[loop])
        else:
            self.ws.write(self.rowsize, 74, 'Empty')
        if self.xl_Text10[loop]:
            self.ws.write(self.rowsize, 75, self.xl_Text10[loop])
        else:
            self.ws.write(self.rowsize, 75, 'Empty')
        if self.xl_Text11[loop]:
            self.ws.write(self.rowsize, 76, self.xl_Text11[loop])
        else:
            self.ws.write(self.rowsize, 76, 'Empty')
        if self.xl_Text12[loop]:
            self.ws.write(self.rowsize, 77, self.xl_Text12[loop])
        else:
            self.ws.write(self.rowsize, 77, 'Empty')
        if self.xl_Text13[loop]:
            self.ws.write(self.rowsize, 78, self.xl_Text13[loop])
        else:
            self.ws.write(self.rowsize, 78, 'Empty')
        if self.xl_Text14[loop]:
            self.ws.write(self.rowsize, 79, self.xl_Text14[loop])
        else:
            self.ws.write(self.rowsize, 79, 'Empty')
        if self.xl_Text15[loop]:
            self.ws.write(self.rowsize, 80, self.xl_Text15[loop])
        else:
            self.ws.write(self.rowsize, 80, 'Empty')
        if self.xl_TextArea1[loop]:
            self.ws.write(self.rowsize, 81, self.xl_TextArea1[loop])
        else:
            self.ws.write(self.rowsize, 81, 'Empty')
        if self.xl_TextArea2[loop]:
            self.ws.write(self.rowsize, 82, self.xl_TextArea2[loop])
        else:
            self.ws.write(self.rowsize, 82, 'Empty')
        if self.xl_TextArea3[loop]:
            self.ws.write(self.rowsize, 83, self.xl_TextArea3[loop])
        else:
            self.ws.write(self.rowsize, 83, 'Empty')
        if self.xl_TextArea4[loop]:
            self.ws.write(self.rowsize, 84, self.xl_TextArea4[loop])
        else:
            self.ws.write(self.rowsize, 84, 'Empty')
        if self.xl_Exception_message[loop]:
            self.ws.write(self.rowsize, 85, self.xl_Exception_message[loop].format(self.OrginalCID))

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # --------------------------------------------------------------------------------------------------------------

        if self.isCreated:
            self.ws.write(self.rowsize, 1, 'Pass', self.style8)
            self.success_case_01 = 'Pass'

        elif self.xl_Exception_message[loop]:
            if self.xl_Exception_message[loop] is not None:
                if self.message:
                    if self.xl_Exception_message[loop] and 'combination (Email1)' in self.message:
                        self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                        self.success_case_02 = 'Pass'

                    elif self.xl_Exception_message[loop] and 'combination (USN)' in self.message:
                        self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                        self.success_case_03 = 'Pass'

                elif self.candidatesavemessage:
                    if self.xl_Exception_message[loop] and 'Please specify' in self.candidatesavemessage:
                        self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                        self.success_case_04 = 'Pass'

                    elif self.xl_Exception_message[loop] and 'Candidate USN' in self.candidatesavemessage:
                        self.ws.write(self.rowsize, 1, 'Pass', self.style8)
                        self.success_case_05 = 'Pass'
                else:
                    self.ws.write(self.rowsize, 1, 'Fail', self.style8)
                    self.failure_case_01 = 'Fail'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
            self.failure_case_02 = 'Fail'
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 2, self.CID)
        self.ws.write(self.rowsize, 3, self.app_dict.get('EventId', None))
        self.ws.write(self.rowsize, 4, self.app_dict.get('EventName', None))
        self.ws.write(self.rowsize, 5, self.app_dict.get('JobId', None))
        self.ws.write(self.rowsize, 6, self.app_dict.get('JobName', None))
        self.ws.write(self.rowsize, 7, self.app_dict.get('ApplicantId', None))
        self.ws.write(self.rowsize, 8, self.test_dict.get('TestId', None))
        self.ws.write(self.rowsize, 9, self.test_dict.get('TestName', None))
        self.ws.write(self.rowsize, 10, self.OrginalCID, self.style7)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------
        if self.xl_Name[loop] == self.personal_details_dict.get('Name'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 11, self.candidatesavemessage, self.style7)
            elif self.xl_Name[loop] is None:
                self.ws.write(self.rowsize, 11, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 11, self.personal_details_dict.get('Name'), self.style14)
        else:
            self.ws.write(self.rowsize, 11, self.personal_details_dict.get('Name'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FirstName[loop] == self.personal_details_dict.get('FirstName'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 12, self.candidatesavemessage, self.style7)
            elif self.xl_FirstName[loop] is None:
                self.ws.write(self.rowsize, 12, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 12, self.personal_details_dict.get('FirstName'), self.style14)
        elif self.personal_details_dict.get('FirstName'):
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('FirstName'), self.style7)
        else:
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('FirstName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_MiddleName[loop] == self.personal_details_dict.get('MiddleName'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 13, self.candidatesavemessage, self.style7)
            elif self.xl_MiddleName[loop] is None:
                self.ws.write(self.rowsize, 13, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 13, self.personal_details_dict.get('MiddleName'), self.style14)
        else:
            self.ws.write(self.rowsize, 13, self.personal_details_dict.get('MiddleName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_LastName[loop] == self.personal_details_dict.get('LastName'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 14, self.candidatesavemessage, self.style7)
            elif self.xl_LastName[loop] is None:
                self.ws.write(self.rowsize, 14, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 14, self.personal_details_dict.get('LastName'), self.style14)
        else:
            self.ws.write(self.rowsize, 14, self.personal_details_dict.get('LastName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if str(self.xl_Mobile1[loop]) == str(self.personal_details_dict.get('Mobile1')):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 15, self.candidatesavemessage, self.style7)
            elif self.xl_Mobile1[loop] is None:
                self.ws.write(self.rowsize, 15, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 15, self.personal_details_dict.get('Mobile1'), self.style14)
        else:
            self.ws.write(self.rowsize, 15, self.personal_details_dict.get('Mobile1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if str(self.xl_PhoneOffice[loop]) == str(self.personal_details_dict.get('PhoneOffice')):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 16, self.candidatesavemessage, self.style7)
            elif self.xl_PhoneOffice[loop] is None:
                self.ws.write(self.rowsize, 16, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 16, self.personal_details_dict.get('PhoneOffice'), self.style14)
        else:
            self.ws.write(self.rowsize, 16, self.personal_details_dict.get('PhoneOffice'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Email1[loop] == self.personal_details_dict.get('Email1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 17, self.candidatesavemessage, self.style7)
            elif self.xl_Email1[loop] is None:
                self.ws.write(self.rowsize, 17, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Email1'), self.style14)
        else:
            self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Email1'), self.style14)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Email2[loop] == self.personal_details_dict.get('Email2'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 18, self.candidatesavemessage, self.style7)
            elif self.xl_Email2[loop] is None:
                self.ws.write(self.rowsize, 18, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Email2'), self.style14)
        else:
            self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Email2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Gender[loop] == self.personal_details_dict.get('Gender'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 19, self.candidatesavemessage, self.style7)
            elif self.xl_Gender[loop] is None:
                self.ws.write(self.rowsize, 19, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 19, self.personal_details_dict.get('Gender'), self.style14)
        else:
            self.ws.write(self.rowsize, 19, self.personal_details_dict.get('Gender'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_MaritalStatus[loop] == self.personal_details_dict.get('MaritalStatus'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 20, self.candidatesavemessage, self.style7)
            elif self.xl_MaritalStatus[loop] is None:
                self.ws.write(self.rowsize, 20, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 20, self.personal_details_dict.get('MaritalStatus'), self.style14)
        else:
            self.ws.write(self.rowsize, 20, self.personal_details_dict.get('MaritalStatus'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_DateOfBirth[loop] == self.personal_details_dict.get('DateOfBirth'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 22, self.candidatesavemessage, self.style7)
            elif self.xl_DateOfBirth[loop] is None:
                self.ws.write(self.rowsize, 21, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 21, self.personal_details_dict.get('DateOfBirth'), self.style14)
        elif self.personal_details_dict.get('DateOfBirth'):
            if '00:00:00' in self.personal_details_dict.get('DateOfBirth'):
                self.ws.write(self.rowsize, 21, self.personal_details_dict.get('DateOfBirth'), self.style7)
        else:
            self.ws.write(self.rowsize, 21, self.personal_details_dict.get('DateOfBirth'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_USN[loop] == self.personal_details_dict.get('USN'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 22, self.candidatesavemessage, self.style7)
            elif self.xl_USN[loop] is None:
                self.ws.write(self.rowsize, 22, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 22, self.personal_details_dict.get('USN'), self.style14)
        else:
            self.ws.write(self.rowsize, 22, self.personal_details_dict.get('USN'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Address1[loop] == self.personal_details_dict.get('Address1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 23, self.candidatesavemessage, self.style7)
            elif self.xl_Address1[loop] is None:
                self.ws.write(self.rowsize, 23, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 23, self.personal_details_dict.get('Address1'), self.style14)
        else:
            self.ws.write(self.rowsize, 23, self.personal_details_dict.get('Address1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Address2[loop] == self.personal_details_dict.get('Address2'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 24, self.candidatesavemessage, self.style7)
            elif self.xl_Address2[loop] is None:
                self.ws.write(self.rowsize, 24, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 24, self.personal_details_dict.get('Address2'), self.style14)
        else:
            self.ws.write(self.rowsize, 24, self.personal_details_dict.get('Address2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalPercentage[loop] == self.final_degree_dict.get('Percentage'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 25, self.candidatesavemessage, self.style7)
            if self.xl_FinalPercentage[loop] is None:
                self.ws.write(self.rowsize, 25, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 25, self.final_degree_dict.get('Percentage'), self.style14)
        else:
            self.ws.write(self.rowsize, 25, self.final_degree_dict.get('Percentage'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalEndYear[loop] == self.final_degree_dict.get('EndYear'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 26, self.candidatesavemessage, self.style7)
            elif self.xl_FinalEndYear[loop] is None:
                self.ws.write(self.rowsize, 26, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 26, self.final_degree_dict.get('EndYear'), self.style14)
        else:
            self.ws.write(self.rowsize, 26, self.final_degree_dict.get('EndYear'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalDegreeId[loop] == self.final_degree_dict.get('DegreeId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 27, self.candidatesavemessage, self.style7)
            elif self.xl_FinalDegreeId[loop] is None:
                self.ws.write(self.rowsize, 27, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 27, self.final_degree_dict.get('DegreeId'), self.style14)
        else:
            self.ws.write(self.rowsize, 27, self.final_degree_dict.get('DegreeId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalCollegeId[loop] == self.final_degree_dict.get('CollegeId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 28, self.candidatesavemessage, self.style7)
            elif self.xl_FinalCollegeId[loop] is None:
                self.ws.write(self.rowsize, 28, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 28, self.final_degree_dict.get('CollegeId'), self.style14)
        else:
            self.ws.write(self.rowsize, 28, self.final_degree_dict.get('CollegeId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_FinalDegreeTypeId[loop] == self.final_degree_dict.get('DegreeTypeId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 29, self.candidatesavemessage, self.style7)
            elif self.xl_FinalDegreeTypeId[loop] is None:
                self.ws.write(self.rowsize, 29, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 29, self.final_degree_dict.get('DegreeTypeId'), self.style14)
        else:
            self.ws.write(self.rowsize, 29, self.final_degree_dict.get('DegreeTypeId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_10thPercentage[loop] == self.tenth_dict.get('Percentage'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 30, self.candidatesavemessage, self.style7)
            elif self.xl_10thPercentage[loop] is None:
                self.ws.write(self.rowsize, 30, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 30, self.tenth_dict.get('Percentage'), self.style14)
        else:
            self.ws.write(self.rowsize, 30, self.tenth_dict.get('Percentage'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_10thEndYear[loop] == self.tenth_dict.get('EndYear'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 31, self.candidatesavemessage, self.style7)
            elif self.xl_10thEndYear[loop] is None:
                self.ws.write(self.rowsize, 31, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 31, self.tenth_dict.get('EndYear'), self.style14)
        else:
            self.ws.write(self.rowsize, 31, self.tenth_dict.get('EndYear'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_12thPercentage[loop] == self.twelfth_dict.get('Percentage'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 32, self.candidatesavemessage, self.style7)
            elif self.xl_12thPercentage[loop] is None:
                self.ws.write(self.rowsize, 32, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 32, self.twelfth_dict.get('Percentage'), self.style14)
        else:
            self.ws.write(self.rowsize, 32, self.twelfth_dict.get('Percentage'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_12thEndYear[loop] == self.twelfth_dict.get('EndYear'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 33, self.candidatesavemessage, self.style7)
            elif self.xl_12thEndYear[loop] is None:
                self.ws.write(self.rowsize, 33, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 33, self.twelfth_dict.get('EndYear'), self.style14)
        else:
            self.ws.write(self.rowsize, 33, self.twelfth_dict.get('EndYear'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_PanNo[loop] == self.personal_details_dict.get('PanNo'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 34, self.candidatesavemessage, self.style7)
            elif self.xl_PanNo[loop] is None:
                self.ws.write(self.rowsize, 34, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 34, self.personal_details_dict.get('PanNo'), self.style14)
        else:
            self.ws.write(self.rowsize, 34, self.personal_details_dict.get('PanNo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_PassportNo[loop] == self.personal_details_dict.get('PassportNo'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 35, self.candidatesavemessage, self.style7)
            elif self.xl_PassportNo[loop] is None:
                self.ws.write(self.rowsize, 35, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 35, self.personal_details_dict.get('PassportNo'), self.style14)
        else:
            self.ws.write(self.rowsize, 35, self.personal_details_dict.get('PassportNo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_CurrentLocationId[loop] == self.personal_details_dict.get('CurrentLocationId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 36, self.candidatesavemessage, self.style7)
            elif self.xl_CurrentLocationId[loop] is None:
                self.ws.write(self.rowsize, 36, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 36, self.personal_details_dict.get('CurrentLocationId'), self.style14)
        else:
            self.ws.write(self.rowsize, 36, self.personal_details_dict.get('CurrentLocationId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TotalExperienceInMonths[loop] == self.personal_details_dict.get('TotalExperienceInYears'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 37, self.candidatesavemessage, self.style7)
            elif self.xl_TotalExperienceInMonths[loop] is None:
                self.ws.write(self.rowsize, 37, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 37,
                              '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears'),
                                             self.personal_details_dict.get('TotalExperienceInMonths'),
                                             '--Converting to Year(s) & Month(s)'), self.style7)
        elif self.personal_details_dict.get('TotalExperienceInYears'):
            self.ws.write(self.rowsize, 37,
                          '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears'),
                                         self.personal_details_dict.get('TotalExperienceInMonths'),
                                         '--Converting to Year(s) & Month(s)'), self.style7)
        else:
            self.ws.write(self.rowsize, 37, self.personal_details_dict.get('TotalExperienceInYears'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Country[loop] == self.personal_details_dict.get('Country'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 38, self.candidatesavemessage, self.style7)
            elif self.xl_Country[loop] is None:
                self.ws.write(self.rowsize, 38, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 38, self.personal_details_dict.get('Country'), self.style14)
        else:
            self.ws.write(self.rowsize, 38, self.personal_details_dict.get('Country'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_HierarchyId[loop] == self.personal_details_dict.get('HierarchyId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 39, self.candidatesavemessage, self.style7)
            elif self.xl_HierarchyId[loop] is None:
                self.ws.write(self.rowsize, 39, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 39, self.personal_details_dict.get('HierarchyId'), self.style14)
        else:
            self.ws.write(self.rowsize, 39, self.personal_details_dict.get('HierarchyId'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Nationality[loop] == self.personal_details_dict.get('Nationality'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 40, self.candidatesavemessage, self.style7)
            elif self.xl_Nationality[loop] is None:
                self.ws.write(self.rowsize, 40, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 40, self.personal_details_dict.get('Nationality'), self.style14)
        else:
            self.ws.write(self.rowsize, 40, self.personal_details_dict.get('Nationality'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Sensitivity[loop] == self.personal_details_dict.get('Sensitivity'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 41, self.candidatesavemessage, self.style7)
            elif self.xl_Sensitivity[loop] is None:
                self.ws.write(self.rowsize, 41, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 41, self.personal_details_dict.get('Sensitivity'), self.style14)
        else:
            self.ws.write(self.rowsize, 41, self.personal_details_dict.get('Sensitivity'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_StatusId[loop] == self.personal_details_dict.get('StatusId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 42, self.candidatesavemessage, self.style7)
            elif self.xl_StatusId[loop] is None:
                self.ws.write(self.rowsize, 42, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 42, self.personal_details_dict.get('StatusId'), self.style14)
        else:
            self.ws.write(self.rowsize, 42, self.personal_details_dict.get('StatusId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_SourceId[loop] == self.source_details_dict.get('SourceId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 43, self.candidatesavemessage, self.style7)
            elif self.xl_SourceId[loop] is None:
                self.ws.write(self.rowsize, 43, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 43, self.source_details_dict.get('SourceId'), self.style14)
        else:
            self.ws.write(self.rowsize, 43, self.source_details_dict.get('SourceId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_CampusId[loop] == self.source_details_dict.get('CampusId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 44, self.candidatesavemessage, self.style7)
            elif self.xl_CampusId[loop] is None:
                self.ws.write(self.rowsize, 44, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 44, self.source_details_dict.get('CampusId'), self.style14)
        else:
            self.ws.write(self.rowsize, 44, self.source_details_dict.get('CampusId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_SourceType[loop] == self.source_details_dict.get('SourceType'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 45, self.candidatesavemessage, self.style7)
            elif self.xl_SourceType[loop] is None:
                self.ws.write(self.rowsize, 45, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 45, self.source_details_dict.get('SourceType'), self.style14)
        elif self.source_details_dict.get('SourceType') == 6:
            self.ws.write(self.rowsize, 45, self.source_details_dict.get('SourceType'), self.style7)
        else:
            self.ws.write(self.rowsize, 45, self.source_details_dict.get('SourceType'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Experience[loop] == self.experience_dict.get('Experience'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 46, self.candidatesavemessage, self.style7)
            elif self.xl_Experience[loop] is None:
                self.ws.write(self.rowsize, 46, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 46, self.experience_dict.get('Experience'), self.style14)
        else:
            self.ws.write(self.rowsize, 46, self.experience_dict.get('Experience'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_EmployerId[loop] == self.experience_dict.get('EmployerId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 47, self.candidatesavemessage, self.style7)
            elif self.xl_EmployerId[loop] is None:
                self.ws.write(self.rowsize, 47, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 47, self.experience_dict.get('EmployerId'), self.style14)
        else:
            self.ws.write(self.rowsize, 47, self.experience_dict.get('EmployerId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_DesignationId[loop] == self.experience_dict.get('DesignationId'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 48, self.candidatesavemessage, self.style7)
            elif self.xl_DesignationId[loop] is None:
                self.ws.write(self.rowsize, 48, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 48, self.experience_dict.get('DesignationId'), self.style14)
        else:
            self.ws.write(self.rowsize, 48, self.experience_dict.get('DesignationId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Expertise[loop] == self.personal_details_dict.get('ExpertiseId1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 49, self.candidatesavemessage, self.style7)
            elif self.xl_Expertise[loop] is None:
                self.ws.write(self.rowsize, 49, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 49, self.personal_details_dict.get('ExpertiseId1'), self.style14)
        else:
            self.ws.write(self.rowsize, 49, self.personal_details_dict.get('ExpertiseId1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_NoticePeriod[loop] == self.personal_details_dict.get('NoticePeriod'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 50, self.candidatesavemessage, self.style7)
            elif self.xl_NoticePeriod[loop] is None:
                self.ws.write(self.rowsize, 50, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 50, self.personal_details_dict.get('NoticePeriod'), self.style14)
        else:
            self.ws.write(self.rowsize, 50, self.personal_details_dict.get('NoticePeriod'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer1[loop] == self.custom_details_dict.get('Integer1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 51, self.candidatesavemessage, self.style7)
            elif self.xl_Integer1[loop] is None:
                self.ws.write(self.rowsize, 51, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 51, self.custom_details_dict.get('Integer1'), self.style14)
        else:
            self.ws.write(self.rowsize, 51, self.custom_details_dict.get('Integer1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer2[loop] == self.custom_details_dict.get('Integer2'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 52, self.candidatesavemessage, self.style7)
            elif self.xl_Integer2[loop] is None:
                self.ws.write(self.rowsize, 52, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer2'), self.style14)
        else:
            self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer3[loop] == self.custom_details_dict.get('Integer3'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 53, self.candidatesavemessage, self.style7)
            elif self.xl_Integer3[loop] is None:
                self.ws.write(self.rowsize, 53, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer3'), self.style14)
        else:
            self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer4[loop] == self.custom_details_dict.get('Integer4'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 54, self.candidatesavemessage, self.style7)
            elif self.xl_Integer4[loop] is None:
                self.ws.write(self.rowsize, 54, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer4'), self.style14)
        else:
            self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer5[loop] == self.custom_details_dict.get('Integer5'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 55, self.candidatesavemessage, self.style7)
            elif self.xl_Integer5[loop] is None:
                self.ws.write(self.rowsize, 55, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer5'), self.style14)
        else:
            self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer5'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer6[loop] == self.custom_details_dict.get('Integer6'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 56, self.candidatesavemessage, self.style7)
            elif self.xl_Integer6[loop] is None:
                self.ws.write(self.rowsize, 56, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer6'), self.style14)
        else:
            self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer6'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer7[loop] == self.custom_details_dict.get('Integer7'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 57, self.candidatesavemessage, self.style7)
            elif self.xl_Integer7[loop] is None:
                self.ws.write(self.rowsize, 57, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer7'), self.style14)
        else:
            self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer7'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer8[loop] == self.custom_details_dict.get('Integer8'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 58, self.candidatesavemessage, self.style7)
            elif self.xl_Integer8[loop] is None:
                self.ws.write(self.rowsize, 58, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer8'), self.style14)
        else:
            self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer8'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer9[loop] == self.custom_details_dict.get('Integer9'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 59, self.candidatesavemessage, self.style7)
            elif self.xl_Integer9[loop] is None:
                self.ws.write(self.rowsize, 59, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer9'), self.style14)
        else:
            self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer9'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer10[loop] == self.custom_details_dict.get('Integer10'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 60, self.candidatesavemessage, self.style7)
            elif self.xl_Integer10[loop] is None:
                self.ws.write(self.rowsize, 60, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Integer10'), self.style14)
        else:
            self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Integer10'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer11[loop] == self.custom_details_dict.get('Integer11'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 61, self.candidatesavemessage, self.style7)
            elif self.xl_Integer11[loop] is None:
                self.ws.write(self.rowsize, 61, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Integer11'), self.style14)
        else:
            self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Integer11'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer12[loop] == self.custom_details_dict.get('Integer12'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 62, self.candidatesavemessage, self.style7)
            elif self.xl_Integer12[loop] is None:
                self.ws.write(self.rowsize, 62, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Integer12'), self.style14)
        else:
            self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Integer12'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer13[loop] == self.custom_details_dict.get('Integer13'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 63, self.candidatesavemessage, self.style7)
            elif self.xl_Integer13[loop] is None:
                self.ws.write(self.rowsize, 63, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Integer13'), self.style14)
        else:
            self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Integer13'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer14[loop] == self.custom_details_dict.get('Integer14'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 64, self.candidatesavemessage, self.style7)
            elif self.xl_Integer14[loop] is None:
                self.ws.write(self.rowsize, 64, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Integer14'), self.style14)
        else:
            self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Integer14'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Integer15[loop] == self.custom_details_dict.get('Integer15'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 65, self.candidatesavemessage, self.style7)
            elif self.xl_Integer15[loop] is None:
                self.ws.write(self.rowsize, 65, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Integer15'), self.style14)
        else:
            self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Integer15'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text1[loop] == self.custom_details_dict.get('Text1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 66, self.candidatesavemessage, self.style7)
            elif self.xl_Text1[loop] is None:
                self.ws.write(self.rowsize, 66, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Text1'), self.style14)
        else:
            self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Text1'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text2[loop] == self.custom_details_dict.get('Text2'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 67, self.candidatesavemessage, self.style7)
            elif self.xl_Text2[loop] is None:
                self.ws.write(self.rowsize, 67, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text2'), self.style14)
        else:
            self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text2'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text3[loop] == self.custom_details_dict.get('Text3'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 68, self.candidatesavemessage, self.style7)
            elif self.xl_Text3[loop] is None:
                self.ws.write(self.rowsize, 68, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text3'), self.style14)
        else:
            self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text4[loop] == self.custom_details_dict.get('Text4'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 69, self.candidatesavemessage, self.style7)
            elif self.xl_Text4[loop] is None:
                self.ws.write(self.rowsize, 69, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text4'), self.style14)
        else:
            self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text5[loop] == self.custom_details_dict.get('Text5'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 70, self.candidatesavemessage, self.style7)
            elif self.xl_Text5[loop] is None:
                self.ws.write(self.rowsize, 70, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text5'), self.style14)
        else:
            self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text5'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text6[loop] == self.custom_details_dict.get('Text6'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 71, self.candidatesavemessage, self.style7)
            elif self.xl_Text6[loop] is None:
                self.ws.write(self.rowsize, 71, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text6'), self.style14)
        else:
            self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text6'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text7[loop] == self.custom_details_dict.get('Text7'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 72, self.candidatesavemessage, self.style7)
            elif self.xl_Text7[loop] is None:
                self.ws.write(self.rowsize, 72, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text7'), self.style14)
        else:
            self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text7'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text8[loop] == self.custom_details_dict.get('Text8'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 73, self.candidatesavemessage, self.style7)
            elif self.xl_Text8[loop] is None:
                self.ws.write(self.rowsize, 73, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text8'), self.style14)
        else:
            self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text8'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text9[loop] == self.custom_details_dict.get('Text9'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 74, self.candidatesavemessage, self.style7)
            elif self.xl_Text9[loop] is None:
                self.ws.write(self.rowsize, 74, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text9'), self.style14)
        else:
            self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text9'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text10[loop] == self.custom_details_dict.get('Text10'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 75, self.candidatesavemessage, self.style7)
            elif self.xl_Text10[loop] is None:
                self.ws.write(self.rowsize, 75, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 75, self.custom_details_dict.get('Text10'), self.style14)
        else:
            self.ws.write(self.rowsize, 75, self.custom_details_dict.get('Text10'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text11[loop] == self.custom_details_dict.get('Text11'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 76, self.candidatesavemessage, self.style7)
            elif self.xl_Text11[loop] is None:
                self.ws.write(self.rowsize, 76, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 76, self.custom_details_dict.get('Text11'), self.style14)
        else:
            self.ws.write(self.rowsize, 76, self.custom_details_dict.get('Text11'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text12[loop] == self.custom_details_dict.get('Text12'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 77, self.candidatesavemessage, self.style7)
            elif self.xl_Text12[loop] is None:
                self.ws.write(self.rowsize, 77, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 77, self.custom_details_dict.get('Text12'), self.style14)
        else:
            self.ws.write(self.rowsize, 77, self.custom_details_dict.get('Text12'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text13[loop] == self.custom_details_dict.get('Text13'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 78, self.candidatesavemessage, self.style7)
            elif self.xl_Text13[loop] is None:
                self.ws.write(self.rowsize, 78, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 78, self.custom_details_dict.get('Text13'), self.style14)
        else:
            self.ws.write(self.rowsize, 78, self.custom_details_dict.get('Text13'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text14[loop] == self.custom_details_dict.get('Text14'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 79, self.candidatesavemessage, self.style7)
            elif self.xl_Text14[loop] is None:
                self.ws.write(self.rowsize, 79, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 79, self.custom_details_dict.get('Text14'), self.style14)
        else:
            self.ws.write(self.rowsize, 79, self.custom_details_dict.get('Text14'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Text15[loop] == self.custom_details_dict.get('Text15'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 80, self.candidatesavemessage, self.style7)
            elif self.xl_Text15[loop] is None:
                self.ws.write(self.rowsize, 80, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 80, self.custom_details_dict.get('Text15'), self.style14)
        else:
            self.ws.write(self.rowsize, 80, self.custom_details_dict.get('Text15'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea1[loop] == self.custom_details_dict.get('TextArea1'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 81, self.candidatesavemessage, self.style7)
            elif self.xl_TextArea1[loop] is None:
                self.ws.write(self.rowsize, 81, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 81, self.custom_details_dict.get('TextArea1'), self.style14)
        else:
            self.ws.write(self.rowsize, 81, self.custom_details_dict.get('TextArea1'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea2[loop] == self.custom_details_dict.get('TextArea2'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 82, self.candidatesavemessage, self.style7)
            elif self.xl_TextArea2[loop] is None:
                self.ws.write(self.rowsize, 82, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 82, self.custom_details_dict.get('TextArea2'), self.style14)
        else:
            self.ws.write(self.rowsize, 82, self.custom_details_dict.get('TextArea2'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea3[loop] == self.custom_details_dict.get('TextArea3'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 83, self.candidatesavemessage, self.style7)
            if self.xl_TextArea3[loop] is None:
                self.ws.write(self.rowsize, 83, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 83, self.custom_details_dict.get('TextArea3'), self.style14)
        else:
            self.ws.write(self.rowsize, 83, self.custom_details_dict.get('TextArea3'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.xl_TextArea4[loop] == self.custom_details_dict.get('TextArea4'):
            if self.candidatesavemessage:
                self.ws.write(self.rowsize, 84, self.candidatesavemessage, self.style7)
            if self.xl_TextArea4[loop] is None:
                self.ws.write(self.rowsize, 84, 'Empty', self.style14)
            else:
                self.ws.write(self.rowsize, 84, self.custom_details_dict.get('TextArea4'), self.style14)
        else:
            self.ws.write(self.rowsize, 84, self.custom_details_dict.get('TextArea4'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if self.message and 'Duplicate candidate found with combination' in self.xl_Exception_message[loop]:
            self.ws.write(self.rowsize, 85, self.message, self.style14)
        elif self.xl_Exception_message[loop] == self.partialmessage:
            self.ws.write(self.rowsize, 85, self.partialmessage, self.style14)
        elif self.xl_Exception_message[loop] == self.candidatesavemessage:
            self.ws.write(self.rowsize, 85, self.candidatesavemessage, self.style14)
        else:
            self.ws.write(self.rowsize, 85, self.message, self.style3)

        # --------------------------------------------------------------------------------------------------------------
        self.rowsize += 1  # Row increment
        Obj.wb_Result.save(output_paths.outputpaths['Candidate_Output_sheet'])

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

    def overall_status(self):
        self.ws.write(0, 0, 'Upload Candidates', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(0, 1, 'Pass', self.style24)
        else:
            self.ws.write(0, 1, 'Fail', self.style25)

        self.ws.write(0, 2, 'Start Time', self.style23)
        self.ws.write(0, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Obj.wb_Result.save(output_paths.outputpaths['Candidate_Output_sheet'])


Obj = UploadCandidate()
Obj.excel_data()
Total_count = len(Obj.xl_Name)
print("Number Of Rows ::", Total_count)
if Obj.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Obj.bulk_create_tag_candidates(looping)
        if Obj.isCreated:  # Always Boolean is true, if it is not mention
            Obj.candidate_get_by_id_details()
            Obj.candidate_educational_details(looping)
            Obj.candidate_experience_details()
            Obj.event_applicants(looping)
        Obj.output_excel(looping)
        Obj.personal_details_dict = {}
        Obj.source_details_dict = {}
        Obj.custom_details_dict = {}
        Obj.final_degree_dict = {}
        Obj.tenth_dict = {}
        Obj.twelfth_dict = {}
        Obj.experience_dict = {}
        Obj.app_dict = {}
        Obj.test_dict = {}
        Obj.isCreated = {}
        Obj.OrginalCID = {}
        Obj.message = {}
        Obj.CID = {}
        Obj.candidatesavemessage = {}
        Obj.applicantDetails = {}
        Obj.partialmessage = {}
        Obj.success_case_01 = {}
        Obj.success_case_02 = {}
        Obj.success_case_03 = {}
        Obj.success_case_04 = {}
        Obj.success_case_05 = {}
        Obj.headers = {}

Obj.overall_status()
