from hpro_automation import (login, input_paths, work_book, output_paths, api)
import xlrd
import datetime
import json
import requests


class UpdateCandidate(login.CRPOLogin, work_book.WorkBook):

    def __init__(self):

        self.start_time = str(datetime.datetime.now())

        super(UpdateCandidate, self).__init__()

        self.Expected_success_cases = list(map(lambda x: 'Pass', range(0, 37)))
        self.Actual_Success_case = []

        self.xl_update_candidate_id = []
        self.xl_update_Name = []
        self.xl_update_FirstName = []
        self.xl_update_MiddleName = []
        self.xl_update_LastName = []
        self.xl_update_TotalExperienceInMonths = []
        self.xl_update_passport = []
        self.xl_update_CurrentCTC = []
        self.xl_update_Mobile1 = []
        self.xl_update_PanNo = []
        self.xl_update_Martial = []
        self.xl_update_DOB = []
        self.xl_update_CurrentLocId = []
        self.xl_update_CurrentLocText = []
        self.xl_update_Address1 = []
        self.xl_update_ExpertiseId1 = []
        self.xl_update_PhoneOffice = []
        self.xl_update_Gender = []
        self.xl_update_TotalExperienceInYears = []
        self.xl_update_Email2 = []
        self.xl_update_Email1 = []
        self.xl_update_CurrencyType = []
        self.xl_update_Sensitivity = []
        self.xl_update_Nationality = []
        self.xl_update_Country = []
        self.xl_update_StatusId = []
        self.xl_update_USN = []
        self.xl_update_AadharNo = []
        self.xl_update_HierarchyId = []
        self.xl_update_DesiredSalaryFrom = []
        self.xl_update_DesiredSalaryTo = []
        self.xl_update_NoticePeriod = []
        self.xl_update_SourceId = []
        self.xl_update_SourceType = []
        self.xl_update_Skill1 = []
        self.xl_update_Skill2 = []
        self.xl_update_LinkedInLink = []
        self.xl_update_FacebookLink = []
        self.xl_update_TwitterLink = []
        self.xl_update_Integer1 = []
        self.xl_update_Integer2 = []
        self.xl_update_Integer3 = []
        self.xl_update_Integer4 = []
        self.xl_update_Integer5 = []
        self.xl_update_Integer6 = []
        self.xl_update_Integer7 = []
        self.xl_update_Integer8 = []
        self.xl_update_Integer9 = []
        self.xl_update_Integer10 = []
        self.xl_update_Integer11 = []
        self.xl_update_Integer12 = []
        self.xl_update_Integer13 = []
        self.xl_update_Integer14 = []
        self.xl_update_Integer15 = []
        self.xl_update_Text1 = []
        self.xl_update_Text2 = []
        self.xl_update_Text3 = []
        self.xl_update_Text4 = []
        self.xl_update_Text5 = []
        self.xl_update_Text6 = []
        self.xl_update_Text7 = []
        self.xl_update_Text8 = []
        self.xl_update_Text9 = []
        self.xl_update_Text10 = []
        self.xl_update_Text11 = []
        self.xl_update_Text12 = []
        self.xl_update_Text13 = []
        self.xl_update_Text14 = []
        self.xl_update_Text15 = []
        self.xl_update_TextArea1 = []
        self.xl_update_TextArea2 = []
        self.xl_update_TextArea3 = []
        self.xl_update_TextArea4 = []
        self.xl_update_TrueFalse1 = []
        self.xl_update_TrueFalse2 = []
        self.xl_update_TrueFalse3 = []
        self.xl_update_TrueFalse4 = []
        self.xl_update_TrueFalse5 = []
        self.xl_update_DateCustomField1 = []
        self.xl_update_DateCustomField2 = []
        self.xl_update_DateCustomField3 = []
        self.xl_update_DateCustomField4 = []
        self.xl_update_DateCustomField5 = []
        self.xl_updated_expected_message = []

        # --------- Dict --------------------
        self.update_personal_details_dict = {}
        self.update_source_details_dict = {}
        self.update_custom_details_dict = {}
        self.update_social_details_dict = {}
        self.update_primary_skills_dict = {}
        self.update_secondary_skills_dict = {}
        self.update_candidate_preference_dict = {}
        self.api_updated_CID = {}
        self.error_description = {}
        self.success_case_01 = {}
        self.success_case_02 = {}
        self.headers = {}

    def excel_headers(self):
        self.main_headers = ['Comparison', 'Status', 'Candidate ID', 'Name', 'FirstName', 'MiddleName', 'LastName',
                             'PassportNo', 'CurrentCTC', 'Mobile1', 'PanNo',	'MaritalStatus',
                             'DateOfBirth', 'CurrentLocationId', 'CurrentLocationText', 'Address1', 'ExpertiseId1',
                             'PhoneOffice', 'Gender', 'Email2', 'Email1', 'CurrencyType',
                             'Sensitivity', 'Nationality', 'Country', 'StatusId', 'USN', 'AadhaarNo', 'HierarchyId',
                             'DesiredSalaryFrom', 'DesiredSalaryTo', 'NoticePeriod', 'SourceId', 'SourceType', 'Skill1',
                             'Skill2', 'LinkedInLink', 'FacebookLink', 'TwitterLink', 'Integer1', 'Integer2',
                             'Integer3', 'Integer4', 'Integer5', 'Integer6', 'Integer7', 'Integer8', 'Integer9',
                             'Integer10', 'Integer11', 'Integer12', 'Integer13', 'Integer14', 'Integer15', 'Text1',
                             'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7', 'Text8', 'Text9', 'Text10', 'Text11',
                             'Text12', 'Text13', 'Text14', 'Text15', 'TextArea1', 'TextArea2', 'TextArea3', 'TextArea4',
                             'DateCustomField1', 'DateCustomField2', 'DateCustomField3', 'DateCustomField4',
                             'DateCustomField5', 'TrueFalse1', 'TrueFalse2', 'TrueFalse3', 'TrueFalse4', 'TrueFalse5',
                             'Expected/Error_message']
        self.headers_with_style2 = ['Comparison', 'Status', 'Candidate Id']
        self.file_headers_col_row()

    def read_excel(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['Update_candi_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0]:
                self.xl_update_candidate_id.append(int(rows[0]))

            if rows[1]:
                self.xl_update_Name.append(rows[1])

            if not rows[2]:
                self.xl_update_FirstName.append(None)
            else:
                self.xl_update_FirstName.append(rows[2])

            if not rows[3]:
                self.xl_update_MiddleName.append(None)
            else:
                self.xl_update_MiddleName.append(rows[3])

            if not rows[4]:
                self.xl_update_LastName.append(None)
            else:
                self.xl_update_LastName.append(rows[4])

            if not rows[5]:
                self.xl_update_TotalExperienceInMonths.append(None)
            else:
                self.xl_update_TotalExperienceInMonths.append(int(rows[5]))

            if not rows[6]:
                self.xl_update_passport.append(None)
            else:
                self.xl_update_passport.append(str(rows[6]))

            if not rows[7]:
                self.xl_update_CurrentCTC.append(None)
            else:
                self.xl_update_CurrentCTC.append(int(rows[7]))

            if not rows[8]:
                self.xl_update_Mobile1.append(None)
            else:
                self.xl_update_Mobile1.append(str(int(rows[8])))

            if not rows[9]:
                self.xl_update_PanNo.append(None)
            else:
                self.xl_update_PanNo.append(str(rows[9]))

            if not rows[10]:
                self.xl_update_Martial.append(None)
            else:
                self.xl_update_Martial.append(int(rows[10]))

            if not rows[11]:
                self.xl_update_DOB.append(None)
            else:
                self.xl_update_DOB.append(str(rows[11]))

            if not rows[12]:
                self.xl_update_CurrentLocId.append(None)
            else:
                self.xl_update_CurrentLocId.append(int(rows[12]))

            if not rows[13]:
                self.xl_update_CurrentLocText.append(None)
            else:
                self.xl_update_CurrentLocText.append(str(rows[13]))

            if not rows[14]:
                self.xl_update_Address1.append(None)
            else:
                self.xl_update_Address1.append(str(rows[14]))

            if not rows[15]:
                self.xl_update_ExpertiseId1.append(None)
            else:
                self.xl_update_ExpertiseId1.append(int(rows[15]))

            if not rows[16]:
                self.xl_update_PhoneOffice.append(None)
            else:
                self.xl_update_PhoneOffice.append(str(int(rows[16])))

            if not rows[17]:
                self.xl_update_Gender.append(None)
            else:
                self.xl_update_Gender.append(int(rows[17]))

            if not rows[18]:
                self.xl_update_TotalExperienceInYears.append(None)
            else:
                self.xl_update_TotalExperienceInYears.append(int(rows[18]))

            if not rows[19]:
                self.xl_update_Email2.append(None)
            else:
                self.xl_update_Email2.append(str(rows[19]))

            if not rows[20]:
                self.xl_update_Email1.append(None)
            else:
                self.xl_update_Email1.append(str(rows[20]))

            if not rows[21]:
                self.xl_update_CurrencyType.append(None)
            else:
                self.xl_update_CurrencyType.append(str(rows[21]))

            if not rows[22]:
                self.xl_update_Sensitivity.append(None)
            else:
                self.xl_update_Sensitivity.append(int(rows[22]))

            if not rows[23]:
                self.xl_update_Nationality.append(None)
            else:
                self.xl_update_Nationality.append(int(rows[23]))

            if not rows[24]:
                self.xl_update_Country.append(None)
            else:
                self.xl_update_Country.append(int(rows[24]))

            if not rows[25]:
                self.xl_update_StatusId.append(None)
            else:
                self.xl_update_StatusId.append(int(rows[25]))

            if not rows[26]:
                self.xl_update_USN.append(None)
            else:
                self.xl_update_USN.append(str(rows[26]))

            if not rows[27]:
                self.xl_update_AadharNo.append(None)
            else:
                self.xl_update_AadharNo.append(str(rows[27]))

            if not rows[28]:
                self.xl_update_HierarchyId.append(None)
            else:
                self.xl_update_HierarchyId.append(int(rows[28]))

            if not rows[29]:
                self.xl_update_DesiredSalaryFrom.append(None)
            else:
                self.xl_update_DesiredSalaryFrom.append(int(rows[29]))

            if not rows[30]:
                self.xl_update_DesiredSalaryTo.append(None)
            else:
                self.xl_update_DesiredSalaryTo.append(int(rows[30]))

            if not rows[31]:
                self.xl_update_NoticePeriod.append(None)
            else:
                self.xl_update_NoticePeriod.append(int(rows[31]))

            if not rows[32]:
                self.xl_update_SourceId.append(None)
            else:
                self.xl_update_SourceId.append(int(rows[32]))

            if not rows[33]:
                self.xl_update_SourceType.append(None)
            else:
                self.xl_update_SourceType.append(int(rows[33]))

            if not rows[34]:
                self.xl_update_Skill1.append(None)
            else:
                self.xl_update_Skill1.append(int(rows[34]))

            if not rows[35]:
                self.xl_update_Skill2.append(None)
            else:
                self.xl_update_Skill2.append(int(rows[35]))

            if not rows[36]:
                self.xl_update_LinkedInLink.append(None)
            else:
                self.xl_update_LinkedInLink.append(str(rows[36]))

            if not rows[37]:
                self.xl_update_FacebookLink.append(None)
            else:
                self.xl_update_FacebookLink.append(str(rows[37]))

            if not rows[38]:
                self.xl_update_TwitterLink.append(None)
            else:
                self.xl_update_TwitterLink.append(str(rows[38]))

            if not rows[39]:
                self.xl_update_Integer1.append(None)
            else:
                self.xl_update_Integer1.append(int(rows[39]))

            if not rows[40]:
                self.xl_update_Integer2.append(None)
            else:
                self.xl_update_Integer2.append(int(rows[40]))

            if not rows[41]:
                self.xl_update_Integer3.append(None)
            else:
                self.xl_update_Integer3.append(int(rows[41]))

            if not rows[42]:
                self.xl_update_Integer4.append(None)
            else:
                self.xl_update_Integer4.append(int(rows[42]))

            if not rows[43]:
                self.xl_update_Integer5.append(None)
            else:
                self.xl_update_Integer5.append(int(rows[43]))

            if not rows[44]:
                self.xl_update_Integer6.append(None)
            else:
                self.xl_update_Integer6.append(int(rows[44]))

            if not rows[45]:
                self.xl_update_Integer7.append(None)
            else:
                self.xl_update_Integer7.append(int(rows[45]))

            if not rows[46]:
                self.xl_update_Integer8.append(None)
            else:
                self.xl_update_Integer8.append(int(rows[46]))

            if not rows[47]:
                self.xl_update_Integer9.append(None)
            else:
                self.xl_update_Integer9.append(int(rows[47]))

            if not rows[48]:
                self.xl_update_Integer10.append(None)
            else:
                self.xl_update_Integer10.append(int(rows[48]))

            if not rows[49]:
                self.xl_update_Integer11.append(None)
            else:
                self.xl_update_Integer11.append(int(rows[49]))

            if not rows[50]:
                self.xl_update_Integer12.append(None)
            else:
                self.xl_update_Integer12.append(int(rows[50]))

            if not rows[51]:
                self.xl_update_Integer13.append(None)
            else:
                self.xl_update_Integer13.append(int(rows[51]))

            if not rows[52]:
                self.xl_update_Integer14.append(None)
            else:
                self.xl_update_Integer14.append(int(rows[52]))

            if not rows[53]:
                self.xl_update_Integer15.append(None)
            else:
                self.xl_update_Integer15.append(int(rows[53]))

            if not rows[54]:
                self.xl_update_Text1.append(None)
            else:
                self.xl_update_Text1.append(rows[54])

            if not rows[55]:
                self.xl_update_Text2.append(None)
            else:
                self.xl_update_Text2.append(rows[55])

            if not rows[56]:
                self.xl_update_Text3.append(None)
            else:
                self.xl_update_Text3.append(rows[56])

            if not rows[57]:
                self.xl_update_Text4.append(None)
            else:
                self.xl_update_Text4.append(rows[57])

            if not rows[58]:
                self.xl_update_Text5.append(None)
            else:
                self.xl_update_Text5.append(rows[58])

            if not rows[59]:
                self.xl_update_Text6.append(None)
            else:
                self.xl_update_Text6.append(rows[59])

            if not rows[60]:
                self.xl_update_Text7.append(None)
            else:
                self.xl_update_Text7.append(rows[60])

            if not rows[61]:
                self.xl_update_Text8.append(None)
            else:
                self.xl_update_Text8.append(rows[61])

            if not rows[62]:
                self.xl_update_Text9.append(None)
            else:
                self.xl_update_Text9.append(rows[62])

            if not rows[63]:
                self.xl_update_Text10.append(None)
            else:
                self.xl_update_Text10.append(rows[63])

            if not rows[64]:
                self.xl_update_Text11.append(None)
            else:
                self.xl_update_Text11.append(rows[64])

            if not rows[65]:
                self.xl_update_Text12.append(None)
            else:
                self.xl_update_Text12.append(rows[65])

            if not rows[66]:
                self.xl_update_Text13.append(None)
            else:
                self.xl_update_Text13.append(rows[66])

            if not rows[67]:
                self.xl_update_Text14.append(None)
            else:
                self.xl_update_Text14.append(rows[67])

            if not rows[68]:
                self.xl_update_Text15.append(None)
            else:
                self.xl_update_Text15.append(rows[68])

            if not rows[69]:
                self.xl_update_TextArea1.append(None)
            else:
                self.xl_update_TextArea1.append(rows[69])

            if not rows[70]:
                self.xl_update_TextArea2.append(None)
            else:
                self.xl_update_TextArea2.append(rows[70])

            if not rows[71]:
                self.xl_update_TextArea3.append(None)
            else:
                self.xl_update_TextArea3.append(rows[71])

            if not rows[72]:
                self.xl_update_TextArea4.append(None)
            else:
                self.xl_update_TextArea4.append(rows[72])

            if not rows[72]:
                self.xl_update_TextArea4.append(None)
            else:
                self.xl_update_TextArea4.append(rows[72])

            if not rows[73]:
                self.xl_update_DateCustomField1.append(None)
            else:
                self.xl_update_DateCustomField1.append(rows[73])

            if not rows[74]:
                self.xl_update_DateCustomField2.append(None)
            else:
                self.xl_update_DateCustomField2.append(rows[74])

            if not rows[75]:
                self.xl_update_DateCustomField3.append(None)
            else:
                self.xl_update_DateCustomField3.append(rows[75])

            if not rows[76]:
                self.xl_update_DateCustomField4.append(None)
            else:
                self.xl_update_DateCustomField4.append(rows[76])

            if not rows[77]:
                self.xl_update_DateCustomField5.append(None)
            else:
                self.xl_update_DateCustomField5.append(rows[77])

            if not rows[78]:
                self.xl_update_TrueFalse1.append(None)
            else:
                self.xl_update_TrueFalse1.append(rows[78])

            if not rows[79]:
                self.xl_update_TrueFalse2.append(None)
            else:
                self.xl_update_TrueFalse2.append(rows[79])

            if not rows[80]:
                self.xl_update_TrueFalse3.append(None)
            else:
                self.xl_update_TrueFalse3.append(rows[80])

            if not rows[81]:
                self.xl_update_TrueFalse4.append(None)
            else:
                self.xl_update_TrueFalse4.append(rows[81])

            if not rows[82]:
                self.xl_update_TrueFalse5.append(None)
            else:
                self.xl_update_TrueFalse5.append(rows[82])

            if not rows[83]:
                self.xl_updated_expected_message.append(None)
            else:
                self.xl_updated_expected_message.append(rows[83])

    def update_candidate(self, loop):

        self.lambda_function('update_candidate_details')
        self.headers['APP-NAME'] = 'crpo'

        if self.xl_update_TrueFalse1[loop] == 'true':
            truefalse1 = True
        else:
            truefalse1 = False
        if self.xl_update_TrueFalse2[loop] == 'true':
            truefalse2 = True
        else:
            truefalse2 = False
        if self.xl_update_TrueFalse3[loop] == 'true':
            truefalse3 = True
        else:
            truefalse3 = False
        if self.xl_update_TrueFalse4[loop] == 'true':
            truefalse4 = True
        else:
            truefalse4 = False
        if self.xl_update_TrueFalse5[loop] == 'true':
            truefalse5 = True
        else:
            truefalse5 = False

        request = {
            "PersonalDetails": {
                "Name": self.xl_update_Name[loop],
                "FirstName": self.xl_update_FirstName[loop],
                "MiddleName": self.xl_update_MiddleName[loop],
                "LastName": self.xl_update_LastName[loop],
                "TotalExperienceInMonths": self.xl_update_TotalExperienceInMonths[loop],
                "PassportNo": self.xl_update_passport[loop],
                "CurrentCTC": self.xl_update_CurrentCTC[loop],
                "PanNo": self.xl_update_PanNo[loop],
                "MaritalStatus": self.xl_update_Martial[loop],
                "DateOfBirth": self.xl_update_DOB[loop],
                "CurrentLocationId": self.xl_update_CurrentLocId[loop],
                "CurrentLocationText": self.xl_update_CurrentLocText[loop],
                "Address1": self.xl_update_Address1[loop],
                "ExpertiseId1": self.xl_update_ExpertiseId1[loop],
                "PhoneOffice": self.xl_update_PhoneOffice[loop],
                "Gender": self.xl_update_Gender[loop],
                "TotalExperienceInYears": self.xl_update_TotalExperienceInYears[loop],
                "Email2": self.xl_update_Email2[loop],
                "Email1": self.xl_update_Email1[loop],
                "CurrencyType": self.xl_update_CurrencyType[loop],
                "Sensitivity": self.xl_update_Sensitivity[loop],
                "Nationality": self.xl_update_Nationality[loop],
                "Country": self.xl_update_Country[loop],
                "StatusId": self.xl_update_StatusId[loop],
                "USN": self.xl_update_USN[loop],
                "AadhaarNo": self.xl_update_AadharNo[loop],
                "HierarchyId": self.xl_update_HierarchyId[loop]
            },
            "PreferenceDetails": {"CurrencyType": self.xl_update_CurrencyType[loop],
                                  "DesiredSalaryFrom": self.xl_update_DesiredSalaryFrom[loop],
                                  "DesiredSalaryTo": self.xl_update_DesiredSalaryTo[loop],
                                  "NoticePeriod": self.xl_update_NoticePeriod[loop]},
            "SourceDetails": {
                "SourceId": self.xl_update_SourceId[loop]
            },
            "SocialDetails": {
                "LinkedInLink": self.xl_update_LinkedInLink[loop],
                "FacebookLink": self.xl_update_FacebookLink[loop],
                "TwitterLink": self.xl_update_TwitterLink[loop]
            },
            "SkillsDetails": {
                "KeySkills": [{
                    "Id": self.xl_update_Skill1[loop]
                }],
                "SecondaryKeySkills": [{
                    "Id": self.xl_update_Skill2[loop]
                }]
            },
            "CustomDetails": {
                "Integer1": self.xl_update_Integer1[loop],
                "Integer2": self.xl_update_Integer2[loop],
                "Integer3": self.xl_update_Integer3[loop],
                "Integer4": self.xl_update_Integer4[loop],
                "Integer5": self.xl_update_Integer5[loop],
                "Integer6": self.xl_update_Integer6[loop],
                "Integer7": self.xl_update_Integer7[loop],
                "Integer8": self.xl_update_Integer8[loop],
                "Integer9": self.xl_update_Integer9[loop],
                "Integer10": self.xl_update_Integer10[loop],
                "Integer11": self.xl_update_Integer11[loop],
                "Integer12": self.xl_update_Integer12[loop],
                "Integer13": self.xl_update_Integer13[loop],
                "Integer14": self.xl_update_Integer14[loop],
                "Integer15": self.xl_update_Integer15[loop],
                "Text1": self.xl_update_Text1[loop],
                "Text2": self.xl_update_Text2[loop],
                "Text3": self.xl_update_Text3[loop],
                "Text4": self.xl_update_Text4[loop],
                "Text5": self.xl_update_Text5[loop],
                "Text6": self.xl_update_Text6[loop],
                "Text7": self.xl_update_Text7[loop],
                "Text8": self.xl_update_Text8[loop],
                "Text9": self.xl_update_Text9[loop],
                "Text10": self.xl_update_Text10[loop],
                "Text11": self.xl_update_Text11[loop],
                "Text12": self.xl_update_Text12[loop],
                "Text13": self.xl_update_Text13[loop],
                "Text14": self.xl_update_Text14[loop],
                "Text15": self.xl_update_Text15[loop],
                "DateCustomField1": self.xl_update_DateCustomField1[loop],
                "DateCustomField2": self.xl_update_DateCustomField2[loop],
                "DateCustomField3": self.xl_update_DateCustomField3[loop],
                "DateCustomField4": self.xl_update_DateCustomField4[loop],
                "DateCustomField5": self.xl_update_DateCustomField5[loop],
                "TextArea1": self.xl_update_TextArea1[loop],
                "TextArea2": self.xl_update_TextArea2[loop],
                "TextArea3": self.xl_update_TextArea3[loop],
                "TextArea4": self.xl_update_TextArea4[loop],
                "TrueFalse1": truefalse1,
                "TrueFalse2": truefalse2,
                "TrueFalse3": truefalse3,
                "TrueFalse4": truefalse4,
                "TrueFalse5": truefalse5
            },
            "CandidateId": self.xl_update_candidate_id[loop]
        }
        update_api = requests.post(api.web_api['update_candidate_details'], headers=self.headers,
                                   data=json.dumps(request, default=str), verify=False)
        print(update_api.headers)
        update_api_response = json.loads(update_api.content)
        print(update_api_response)
        self.api_updated_CID = update_api_response.get('CandidateId')
        status = update_api_response['status']
        if status == 'KO':
            error = update_api_response.get('error')
            self.error_description = error.get('errorDescription')

    def mobile_update(self, loop):

        request = {"CandidateId": self.xl_update_candidate_id[loop],
                   "UpdateUserCandidate": {"Mobile1": self.xl_update_Mobile1[loop]}}
        mobile_update_api = requests.post(api.web_api['update_candidate_details'], headers=self.headers,
                                          data=json.dumps(request, default=str), verify=False)
        print(mobile_update_api.headers)
        mobile_update_api_response = json.loads(mobile_update_api.content)
        print(mobile_update_api_response)

    def candidate_get_by_id_details(self, loop):

        self.lambda_function('CandidateGetbyId')
        self.headers['APP-NAME'] = 'crpo'

        get_candidate_details = requests.post(api.web_api['CandidateGetbyId'].format(self.xl_update_candidate_id[loop]),
                                              headers=self.headers)
        print(get_candidate_details.headers)
        candidate_details = json.loads(get_candidate_details.content)
        candidate_dict = candidate_details['Candidate']
        self.update_personal_details_dict = candidate_dict['PersonalDetails']
        self.update_source_details_dict = candidate_dict['SourceDetails']
        self.update_custom_details_dict = candidate_dict['CustomDetails']
        self.update_social_details_dict = candidate_dict['SocialDetails']
        self.update_candidate_preference_dict = candidate_dict['CandidatePreference']

        primary_skills_dict = candidate_dict['PrimarySkills']
        for i in primary_skills_dict:
            if self.xl_update_Skill1[loop] == i.get('Id'):
                self.update_primary_skills_dict = i['Id']

        secondary_skills_dict = candidate_dict['SecondrySkills']
        for j in secondary_skills_dict:
            if self.xl_update_Skill2[loop] == j.get('Id'):
                self.update_secondary_skills_dict = j['Id']

    def out_file(self, loop):

        # ----------------------- Input Data ----------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.style4)

        self.ws.write(self.rowsize, 2, self.xl_update_candidate_id[loop])

        self.ws.write(self.rowsize, 3, self.xl_update_Name[loop])
        self.ws.write(self.rowsize, 4, self.xl_update_FirstName[loop] if self.xl_update_FirstName[loop] else 'Empty')
        self.ws.write(self.rowsize, 5, self.xl_update_MiddleName[loop] if self.xl_update_MiddleName[loop] else 'Empty')
        self.ws.write(self.rowsize, 6, self.xl_update_LastName[loop] if self.xl_update_LastName[loop] else 'Empty')

        # self.ws.write(self.rowsize, 7,
        #               self.xl_update_TotalExperienceInMonths[loop] if self.xl_update_TotalExperienceInMonths[loop]
        #               else 'Empty')
        self.ws.write(self.rowsize, 7, self.xl_update_passport[loop] if self.xl_update_passport[loop] else 'Empty')
        self.ws.write(self.rowsize, 8, self.xl_update_CurrentCTC[loop] if self.xl_update_CurrentCTC[loop] else 'Empty')
        self.ws.write(self.rowsize, 9, self.xl_update_Mobile1[loop] if self.xl_update_Mobile1[loop] else 'Empty')
        self.ws.write(self.rowsize, 10, self.xl_update_PanNo[loop] if self.xl_update_PanNo[loop] else 'Empty')
        self.ws.write(self.rowsize, 11, self.xl_update_Martial[loop] if self.xl_update_Martial[loop] else 'Empty')
        self.ws.write(self.rowsize, 12, self.xl_update_DOB[loop] if self.xl_update_DOB[loop] else 'Empty')
        self.ws.write(self.rowsize, 13,
                      self.xl_update_CurrentLocId[loop] if self.xl_update_CurrentLocId[loop] else 'Empty')
        self.ws.write(self.rowsize, 14,
                      self.xl_update_CurrentLocText[loop] if self.xl_update_CurrentLocText[loop] else 'Empty')
        self.ws.write(self.rowsize, 15, self.xl_update_Address1[loop] if self.xl_update_Address1[loop] else 'Empty')
        self.ws.write(self.rowsize, 16,
                      self.xl_update_ExpertiseId1[loop] if self.xl_update_ExpertiseId1[loop] else 'Empty')
        self.ws.write(self.rowsize, 17,
                      self.xl_update_PhoneOffice[loop] if self.xl_update_PhoneOffice[loop] else 'Empty')
        self.ws.write(self.rowsize, 18, self.xl_update_Gender[loop] if self.xl_update_Gender[loop] else 'Empty')

        self.ws.write(self.rowsize, 19, self.xl_update_Email1[loop] if self.xl_update_Email1[loop] else 'Empty')

        self.ws.write(self.rowsize, 20, self.xl_update_Email2[loop] if self.xl_update_Email2[loop] else 'Empty')

        self.ws.write(self.rowsize, 21,
                      self.xl_update_CurrencyType[loop] if self.xl_update_CurrencyType[loop] else 'Empty')

        self.ws.write(self.rowsize, 22,
                      self.xl_update_Sensitivity[loop] if self.xl_update_Sensitivity[loop] else 'Empty')

        self.ws.write(self.rowsize, 23,
                      self.xl_update_Nationality[loop] if self.xl_update_Nationality[loop] else 'Empty')

        self.ws.write(self.rowsize, 24, self.xl_update_Country[loop] if self.xl_update_Country[loop] else 'Empty')

        self.ws.write(self.rowsize, 25, self.xl_update_StatusId[loop] if self.xl_update_StatusId[loop] else 'Empty')

        self.ws.write(self.rowsize, 26, self.xl_update_USN[loop] if self.xl_update_USN[loop] else 'Empty')

        self.ws.write(self.rowsize, 27, self.xl_update_AadharNo[loop] if self.xl_update_AadharNo[loop] else 'Empty')

        self.ws.write(self.rowsize, 28,
                      self.xl_update_HierarchyId[loop] if self.xl_update_HierarchyId[loop] else 'Empty')

        self.ws.write(self.rowsize, 29,
                      self.xl_update_DesiredSalaryFrom[loop] if self.xl_update_DesiredSalaryFrom[loop] else 'Empty')

        self.ws.write(self.rowsize, 30,
                      self.xl_update_DesiredSalaryTo[loop] if self.xl_update_DesiredSalaryTo[loop] else 'Empty')

        self.ws.write(self.rowsize, 31,
                      self.xl_update_NoticePeriod[loop] if self.xl_update_NoticePeriod[loop] else 'Empty')

        self.ws.write(self.rowsize, 32, self.xl_update_SourceId[loop] if self.xl_update_SourceId[loop] else 'Empty')
        self.ws.write(self.rowsize, 33, self.xl_update_SourceType[loop] if self.xl_update_SourceType[loop] else 'Empty')
        self.ws.write(self.rowsize, 34, self.xl_update_Skill1[loop] if self.xl_update_Skill1[loop] else 'Empty')
        self.ws.write(self.rowsize, 35, self.xl_update_Skill2[loop] if self.xl_update_Skill2[loop] else 'Empty')

        self.ws.write(self.rowsize, 36,
                      self.xl_update_LinkedInLink[loop] if self.xl_update_LinkedInLink[loop] else 'Empty')

        self.ws.write(self.rowsize, 37,
                      self.xl_update_FacebookLink[loop] if self.xl_update_FacebookLink[loop] else 'Empty')

        self.ws.write(self.rowsize, 38,
                      self.xl_update_TwitterLink[loop] if self.xl_update_TwitterLink[loop] else 'Empty')

        self.ws.write(self.rowsize, 39, self.xl_update_Integer1[loop] if self.xl_update_Integer1[loop] else 'Empty')
        self.ws.write(self.rowsize, 40, self.xl_update_Integer2[loop] if self.xl_update_Integer2[loop] else 'Empty')
        self.ws.write(self.rowsize, 41, self.xl_update_Integer3[loop] if self.xl_update_Integer3[loop] else 'Empty')
        self.ws.write(self.rowsize, 42, self.xl_update_Integer4[loop] if self.xl_update_Integer4[loop] else 'Empty')
        self.ws.write(self.rowsize, 43, self.xl_update_Integer5[loop] if self.xl_update_Integer5[loop] else 'Empty')
        self.ws.write(self.rowsize, 44, self.xl_update_Integer6[loop] if self.xl_update_Integer6[loop] else 'Empty')
        self.ws.write(self.rowsize, 45, self.xl_update_Integer7[loop] if self.xl_update_Integer7[loop] else 'Empty')
        self.ws.write(self.rowsize, 46, self.xl_update_Integer8[loop] if self.xl_update_Integer8[loop] else 'Empty')
        self.ws.write(self.rowsize, 47, self.xl_update_Integer9[loop] if self.xl_update_Integer9[loop] else 'Empty')
        self.ws.write(self.rowsize, 48, self.xl_update_Integer10[loop] if self.xl_update_Integer10[loop] else 'Empty')
        self.ws.write(self.rowsize, 49, self.xl_update_Integer11[loop] if self.xl_update_Integer11[loop] else 'Empty')
        self.ws.write(self.rowsize, 50, self.xl_update_Integer12[loop] if self.xl_update_Integer12[loop] else 'Empty')
        self.ws.write(self.rowsize, 51, self.xl_update_Integer13[loop] if self.xl_update_Integer13[loop] else 'Empty')
        self.ws.write(self.rowsize, 52, self.xl_update_Integer14[loop] if self.xl_update_Integer14[loop] else 'Empty')
        self.ws.write(self.rowsize, 53, self.xl_update_Integer15[loop] if self.xl_update_Integer15[loop] else 'Empty')

        self.ws.write(self.rowsize, 54, self.xl_update_Text1[loop] if self.xl_update_Text1[loop] else 'Empty')
        self.ws.write(self.rowsize, 55, self.xl_update_Text2[loop] if self.xl_update_Text2[loop] else 'Empty')
        self.ws.write(self.rowsize, 56, self.xl_update_Text3[loop] if self.xl_update_Text3[loop] else 'Empty')
        self.ws.write(self.rowsize, 57, self.xl_update_Text4[loop] if self.xl_update_Text4[loop] else 'Empty')
        self.ws.write(self.rowsize, 58, self.xl_update_Text5[loop] if self.xl_update_Text5[loop] else 'Empty')
        self.ws.write(self.rowsize, 59, self.xl_update_Text6[loop] if self.xl_update_Text6[loop] else 'Empty')
        self.ws.write(self.rowsize, 60, self.xl_update_Text7[loop] if self.xl_update_Text7[loop] else 'Empty')
        self.ws.write(self.rowsize, 61, self.xl_update_Text8[loop] if self.xl_update_Text8[loop] else 'Empty')
        self.ws.write(self.rowsize, 62, self.xl_update_Text9[loop] if self.xl_update_Text9[loop] else 'Empty')
        self.ws.write(self.rowsize, 63, self.xl_update_Text10[loop] if self.xl_update_Text10[loop] else 'Empty')
        self.ws.write(self.rowsize, 64, self.xl_update_Text11[loop] if self.xl_update_Text11[loop] else 'Empty')
        self.ws.write(self.rowsize, 65, self.xl_update_Text12[loop] if self.xl_update_Text12[loop] else 'Empty')
        self.ws.write(self.rowsize, 66, self.xl_update_Text13[loop] if self.xl_update_Text13[loop] else 'Empty')
        self.ws.write(self.rowsize, 67, self.xl_update_Text14[loop] if self.xl_update_Text14[loop] else 'Empty')
        self.ws.write(self.rowsize, 68, self.xl_update_Text15[loop] if self.xl_update_Text15[loop] else 'Empty')

        self.ws.write(self.rowsize, 69, self.xl_update_TextArea1[loop] if self.xl_update_TextArea1[loop] else 'Empty')
        self.ws.write(self.rowsize, 70, self.xl_update_TextArea2[loop] if self.xl_update_TextArea2[loop] else 'Empty')
        self.ws.write(self.rowsize, 71, self.xl_update_TextArea3[loop] if self.xl_update_TextArea3[loop] else 'Empty')
        self.ws.write(self.rowsize, 72, self.xl_update_TextArea4[loop] if self.xl_update_TextArea4[loop] else 'Empty')

        self.ws.write(self.rowsize, 73,
                      self.xl_update_DateCustomField1[loop] if self.xl_update_DateCustomField1[loop] else 'Empty')
        self.ws.write(self.rowsize, 74,
                      self.xl_update_DateCustomField2[loop] if self.xl_update_DateCustomField2[loop] else 'Empty')
        self.ws.write(self.rowsize, 75,
                      self.xl_update_DateCustomField3[loop] if self.xl_update_DateCustomField3[loop] else 'Empty')
        self.ws.write(self.rowsize, 76,
                      self.xl_update_DateCustomField4[loop] if self.xl_update_DateCustomField4[loop] else 'Empty')
        self.ws.write(self.rowsize, 77,
                      self.xl_update_DateCustomField5[loop] if self.xl_update_DateCustomField5[loop] else 'Empty')

        self.ws.write(self.rowsize, 78, self.xl_update_TrueFalse1[loop] if self.xl_update_TrueFalse1[loop] else 'Empty')
        self.ws.write(self.rowsize, 79, self.xl_update_TrueFalse2[loop] if self.xl_update_TrueFalse2[loop] else 'Empty')
        self.ws.write(self.rowsize, 80, self.xl_update_TrueFalse3[loop] if self.xl_update_TrueFalse3[loop] else 'Empty')
        self.ws.write(self.rowsize, 81, self.xl_update_TrueFalse4[loop] if self.xl_update_TrueFalse4[loop] else 'Empty')
        self.ws.write(self.rowsize, 82, self.xl_update_TrueFalse5[loop] if self.xl_update_TrueFalse5[loop] else 'Empty')
        # self.ws.write(self.rowsize, 83,
        #               self.xl_update_TotalExperienceInYears[loop] if self.xl_update_TotalExperienceInYears[loop]
        #               else 'Empty')
        self.ws.write(self.rowsize, 83, self.xl_updated_expected_message[loop])

        self.rowsize += 1

        # -------- Output ------------
        self.ws.write(self.rowsize, self.col, 'Output', self.style5)
        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------
        if self.api_updated_CID:
            self.ws.write(self.rowsize, 1, 'Pass', self.style26)
            self.success_case_01 = 'Pass'
        elif self.error_description:
            if self.xl_updated_expected_message[loop] == self.error_description:
                self.ws.write(self.rowsize, 1, 'Pass', self.style26)
                self.success_case_02 = 'Pass'
        else:
            self.ws.write(self.rowsize, 1, 'Fail', self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if not self.error_description:
            if self.xl_update_candidate_id[loop] == self.api_updated_CID:
                if self.xl_update_candidate_id[loop] is None:
                    self.ws.write(self.rowsize, 2, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 2, self.update_personal_details_dict.get('CandidateId'), self.style14)
            else:
                self.ws.write(self.rowsize, 2, self.update_personal_details_dict.get('CandidateId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Name[loop] == self.update_personal_details_dict.get('Name'):
                if self.xl_update_Name[loop] is None:
                    self.ws.write(self.rowsize, 3, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 3, self.update_personal_details_dict.get('Name'), self.style14)
            else:
                self.ws.write(self.rowsize, 3, self.update_personal_details_dict.get('Name'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_FirstName[loop] == self.update_personal_details_dict.get('FirstName'):
                if self.xl_update_FirstName[loop] is None:
                    self.ws.write(self.rowsize, 4, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 4, self.update_personal_details_dict.get('FirstName'), self.style14)
            elif not self.xl_update_FirstName[loop]:
                self.ws.write(self.rowsize, 4, self.update_personal_details_dict.get('FirstName'), self.style7)
            else:
                self.ws.write(self.rowsize, 4, self.update_personal_details_dict.get('FirstName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_MiddleName[loop] == self.update_personal_details_dict.get('MiddleName'):
                if self.xl_update_MiddleName[loop] is None:
                    self.ws.write(self.rowsize, 5, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 5, self.update_personal_details_dict.get('MiddleName'), self.style14)
            else:
                self.ws.write(self.rowsize, 5, self.update_personal_details_dict.get('MiddleName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_LastName[loop] == self.update_personal_details_dict.get('LastName'):
                if self.xl_update_LastName[loop] is None:
                    self.ws.write(self.rowsize, 6, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 6, self.update_personal_details_dict.get('LastName'), self.style14)
            else:
                self.ws.write(self.rowsize, 6, self.update_personal_details_dict.get('LastName'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        # if not self.error_description:
        #     if self.xl_update_TotalExperienceInMonths[loop] == self.update_personal_details_dict\
        #             .get('TotalExperienceInMonths'):
        #         if self.xl_update_TotalExperienceInMonths[loop] is None:
        #             self.ws.write(self.rowsize, 7, 'Empty', self.style14)
        #         else:
        #             self.ws.write(self.rowsize, 7,
        #                           self.update_personal_details_dict.get('TotalExperienceInMonths'), self.style14)
        #     else:
        #         self.ws.write(self.rowsize, 7,
        #                       self.update_personal_details_dict.get('TotalExperienceInMonths'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_passport[loop] == self.update_personal_details_dict.get('PassportNo'):
                if self.xl_update_passport[loop] is None:
                    self.ws.write(self.rowsize, 7, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 7, self.update_personal_details_dict.get('PassportNo'), self.style14)
            else:
                self.ws.write(self.rowsize, 7, self.update_personal_details_dict.get('PassportNo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_CurrentCTC[loop] == int(self.update_personal_details_dict.get('CurrentCTC')):
                if self.xl_update_CurrentCTC[loop] is None:
                    self.ws.write(self.rowsize, 8, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 8,
                                  int(self.update_personal_details_dict.get('CurrentCTC')), self.style14)
            else:
                self.ws.write(self.rowsize, 8, self.update_personal_details_dict.get('CurrentCTC'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if str(self.xl_update_Mobile1[loop]) == str(self.update_personal_details_dict.get('Mobile1')):
                if self.xl_update_Mobile1[loop] is None:
                    self.ws.write(self.rowsize, 9, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 9, self.update_personal_details_dict.get('Mobile1'), self.style14)
            else:
                self.ws.write(self.rowsize, 9, self.update_personal_details_dict.get('Mobile1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_PanNo[loop] == self.update_personal_details_dict.get('PanNo'):
                if self.xl_update_PanNo[loop] is None:
                    self.ws.write(self.rowsize, 10, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 10, self.update_personal_details_dict.get('PanNo'), self.style14)
            else:
                self.ws.write(self.rowsize, 10, self.update_personal_details_dict.get('PanNo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Martial[loop] == self.update_personal_details_dict.get('MaritalStatus'):
                if self.xl_update_Martial[loop] is None:
                    self.ws.write(self.rowsize, 11, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 11,
                                  self.update_personal_details_dict.get('MaritalStatus'), self.style14)
            else:
                self.ws.write(self.rowsize, 11, self.update_personal_details_dict.get('MaritalStatus'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DOB[loop] == self.update_personal_details_dict.get('DateOfBirth'):
                if self.xl_update_DOB[loop] is None:
                    self.ws.write(self.rowsize, 12, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 12, self.update_personal_details_dict.get('DateOfBirth'), self.style14)
            elif self.update_personal_details_dict.get('DateOfBirth'):
                if '00:00:00' in self.update_personal_details_dict.get('DateOfBirth'):
                    self.ws.write(self.rowsize, 12, self.update_personal_details_dict.get('DateOfBirth'), self.style7)
            else:
                self.ws.write(self.rowsize, 12, self.update_personal_details_dict.get('DateOfBirth'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_CurrentLocId[loop] == self.update_personal_details_dict.get('CurrentLocationId'):
                if self.xl_update_CurrentLocId[loop] is None:
                    self.ws.write(self.rowsize, 13, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 13,
                                  self.update_personal_details_dict.get('CurrentLocationId'), self.style14)
            else:
                self.ws.write(self.rowsize, 13, self.update_personal_details_dict.get('CurrentLocationId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_CurrentLocText[loop] == self.update_personal_details_dict.get('CurrentLocationText'):
                if self.xl_update_CurrentLocText[loop] is None:
                    self.ws.write(self.rowsize, 14, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 14,
                                  self.update_personal_details_dict.get('CurrentLocationText'), self.style14)
            else:
                self.ws.write(self.rowsize, 14,
                              self.update_personal_details_dict.get('CurrentLocationText'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Address1[loop] == self.update_personal_details_dict.get('Address1'):
                if self.xl_update_Address1[loop] is None:
                    self.ws.write(self.rowsize, 15, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 15, self.update_personal_details_dict.get('Address1'), self.style14)
            else:
                self.ws.write(self.rowsize, 15, self.update_personal_details_dict.get('Address1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_ExpertiseId1[loop] == self.update_personal_details_dict.get('ExpertiseId1'):
                if self.xl_update_ExpertiseId1[loop] is None:
                    self.ws.write(self.rowsize, 16, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 16, self.update_personal_details_dict.get('ExpertiseId1'), self.style14)
            else:
                self.ws.write(self.rowsize, 16, self.update_personal_details_dict.get('ExpertiseId1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if str(self.xl_update_PhoneOffice[loop]) == str(self.update_personal_details_dict.get('PhoneOffice')):
                if self.xl_update_PhoneOffice[loop] is None:
                    self.ws.write(self.rowsize, 17, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 17, self.update_personal_details_dict.get('PhoneOffice'), self.style14)
            else:
                self.ws.write(self.rowsize, 17, self.update_personal_details_dict.get('PhoneOffice'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Gender[loop] == self.update_personal_details_dict.get('Gender'):
                if self.xl_update_Gender[loop] is None:
                    self.ws.write(self.rowsize, 18, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 18, self.update_personal_details_dict.get('Gender'), self.style14)
            else:
                self.ws.write(self.rowsize, 18, self.update_personal_details_dict.get('Gender'), self.style3)
        # --------------------------------------------------------------------------------------------------------------

        if not self.error_description:
            if self.xl_update_Email1[loop] == self.update_personal_details_dict.get('Email1'):
                if self.xl_update_Email1[loop] is None:
                    self.ws.write(self.rowsize, 19, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 19, self.update_personal_details_dict.get('Email1'), self.style14)
            else:
                self.ws.write(self.rowsize, 19, self.update_personal_details_dict.get('Email1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Email2[loop] == self.update_personal_details_dict.get('Email2'):
                if self.xl_update_Email2[loop] is None:
                    self.ws.write(self.rowsize, 20, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 20, self.update_personal_details_dict.get('Email2'), self.style14)
            else:
                self.ws.write(self.rowsize, 20, self.update_personal_details_dict.get('Email2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_CurrencyType[loop] == self.update_personal_details_dict.get('CurrencyType'):
                if self.xl_update_CurrencyType[loop] is None:
                    self.ws.write(self.rowsize, 21, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 21, self.update_personal_details_dict.get('CurrencyType'), self.style14)
            else:
                self.ws.write(self.rowsize, 21, self.update_personal_details_dict.get('CurrencyType'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Sensitivity[loop] == self.update_personal_details_dict.get('Sensitivity'):
                if self.xl_update_Sensitivity[loop] is None:
                    self.ws.write(self.rowsize, 22, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 22, self.update_personal_details_dict.get('Sensitivity'), self.style14)
            else:
                self.ws.write(self.rowsize, 22, self.update_personal_details_dict.get('Sensitivity'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Nationality[loop] == self.update_personal_details_dict.get('Nationality'):
                if self.xl_update_Nationality[loop] is None:
                    self.ws.write(self.rowsize, 23, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 23, self.update_personal_details_dict.get('Nationality'), self.style14)
            else:
                self.ws.write(self.rowsize, 23, self.update_personal_details_dict.get('Nationality'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Country[loop] == self.update_personal_details_dict.get('Country'):
                if self.xl_update_Country[loop] is None:
                    self.ws.write(self.rowsize, 24, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 24, self.update_personal_details_dict.get('Country'), self.style14)
            else:
                self.ws.write(self.rowsize, 24, self.update_personal_details_dict.get('Country'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_StatusId[loop] == self.update_personal_details_dict.get('StatusId'):
                if self.xl_update_StatusId[loop] is None:
                    self.ws.write(self.rowsize, 25, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 25, self.update_personal_details_dict.get('StatusId'), self.style14)
            else:
                self.ws.write(self.rowsize, 25, self.update_personal_details_dict.get('StatusId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_USN[loop] == self.update_personal_details_dict.get('USN'):
                if self.xl_update_USN[loop] is None:
                    self.ws.write(self.rowsize, 26, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 26, self.update_personal_details_dict.get('USN'), self.style14)
            else:
                self.ws.write(self.rowsize, 26, self.update_personal_details_dict.get('USN'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_AadharNo[loop] == self.update_personal_details_dict.get('AadhaarNo'):
                if self.xl_update_AadharNo[loop] is None:
                    self.ws.write(self.rowsize, 27, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 27, self.update_personal_details_dict.get('AadhaarNo'), self.style14)
            else:
                self.ws.write(self.rowsize, 27, self.update_personal_details_dict.get('AadhaarNo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_HierarchyId[loop] == self.update_personal_details_dict.get('HierarchyId'):
                if self.xl_update_HierarchyId[loop] is None:
                    self.ws.write(self.rowsize, 28, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 28, self.update_personal_details_dict.get('HierarchyId'), self.style14)
            else:
                self.ws.write(self.rowsize, 28, self.update_personal_details_dict.get('HierarchyId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DesiredSalaryFrom[loop] == self.update_candidate_preference_dict.get('DesiredSalaryFrom'):
                if self.xl_update_DesiredSalaryFrom[loop] is None:
                    self.ws.write(self.rowsize, 29, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 29,
                                  self.update_candidate_preference_dict.get('DesiredSalaryFrom'), self.style14)
            else:
                self.ws.write(self.rowsize, 29,
                              self.update_candidate_preference_dict.get('DesiredSalaryFrom'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DesiredSalaryTo[loop] == self.update_candidate_preference_dict.get('DesiredSalaryTo'):
                if self.xl_update_DesiredSalaryTo[loop] is None:
                    self.ws.write(self.rowsize, 30, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 30,
                                  self.update_candidate_preference_dict.get('DesiredSalaryTo'), self.style14)
            else:
                self.ws.write(self.rowsize, 30,
                              self.update_candidate_preference_dict.get('DesiredSalaryTo'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_NoticePeriod[loop] == self.update_candidate_preference_dict.get('NoticePeriod'):
                if self.xl_update_NoticePeriod[loop] is None:
                    self.ws.write(self.rowsize, 31, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 31,
                                  self.update_candidate_preference_dict.get('NoticePeriod'), self.style14)
            else:
                self.ws.write(self.rowsize, 31, self.update_candidate_preference_dict.get('NoticePeriod'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_SourceId[loop] == self.update_source_details_dict.get('SourceId'):
                if self.xl_update_SourceId[loop] is None:
                    self.ws.write(self.rowsize, 32, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 32, self.update_source_details_dict.get('SourceId'), self.style14)
            else:
                self.ws.write(self.rowsize, 32, self.update_source_details_dict.get('SourceId'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_SourceType[loop] == self.update_source_details_dict.get('SourceType'):
                if self.xl_update_SourceType[loop] is None:
                    self.ws.write(self.rowsize, 33, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 33, self.update_source_details_dict.get('SourceType'), self.style14)
            elif self.update_source_details_dict.get('SourceType') == 6:
                self.ws.write(self.rowsize, 33, self.update_source_details_dict.get('SourceType'), self.style7)
            else:
                self.ws.write(self.rowsize, 33, self.update_source_details_dict.get('SourceType'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.update_primary_skills_dict:
                if self.xl_update_Skill1[loop] == self.update_primary_skills_dict:
                    if self.xl_update_Skill1[loop] is None:
                        self.ws.write(self.rowsize, 34, 'Empty', self.style14)
                    else:
                        self.ws.write(self.rowsize, 34, self.update_primary_skills_dict, self.style14)
                else:
                    self.ws.write(self.rowsize, 34, self.update_primary_skills_dict, self.style3)
            elif self.xl_update_Skill1[loop] is None:
                self.ws.write(self.rowsize, 34, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.update_secondary_skills_dict:
                if self.xl_update_Skill2[loop] == self.update_secondary_skills_dict:
                    if self.xl_update_Skill2[loop] is None:
                        self.ws.write(self.rowsize, 35, 'Empty', self.style14)
                    else:
                        self.ws.write(self.rowsize, 35, self.update_secondary_skills_dict, self.style14)
                else:
                    self.ws.write(self.rowsize, 35, self.update_secondary_skills_dict, self.style3)
            elif self.xl_update_Skill2[loop] is None:
                self.ws.write(self.rowsize, 35, 'Empty', self.style14)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_LinkedInLink[loop] == self.update_social_details_dict.get('LinkedInLink'):
                if self.xl_update_LinkedInLink[loop] is None:
                    self.ws.write(self.rowsize, 36, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 36, self.update_social_details_dict.get('LinkedInLink'), self.style14)
            elif self.update_social_details_dict.get('LinkedInLink') == 6:
                self.ws.write(self.rowsize, 36, self.update_social_details_dict.get('LinkedInLink'), self.style7)
            else:
                self.ws.write(self.rowsize, 36, self.update_social_details_dict.get('LinkedInLink'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_FacebookLink[loop] == self.update_social_details_dict.get('FacebookLink'):
                if self.xl_update_FacebookLink[loop] is None:
                    self.ws.write(self.rowsize, 37, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 37, self.update_social_details_dict.get('FacebookLink'), self.style14)
            elif self.update_social_details_dict.get('FacebookLink') == 6:
                self.ws.write(self.rowsize, 37, self.update_social_details_dict.get('FacebookLink'), self.style7)
            else:
                self.ws.write(self.rowsize, 37, self.update_social_details_dict.get('FacebookLink'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TwitterLink[loop] == self.update_social_details_dict.get('TwitterLink'):
                if self.xl_update_TwitterLink[loop] is None:
                    self.ws.write(self.rowsize, 38, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 38, self.update_social_details_dict.get('TwitterLink'), self.style14)
            elif self.update_social_details_dict.get('TwitterLink') == 6:
                self.ws.write(self.rowsize, 38, self.update_social_details_dict.get('TwitterLink'), self.style7)
            else:
                self.ws.write(self.rowsize, 38, self.update_social_details_dict.get('TwitterLink'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer1[loop] == self.update_custom_details_dict.get('Integer1'):
                if self.xl_update_Integer1[loop] is None:
                    self.ws.write(self.rowsize, 39, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 39, self.update_custom_details_dict.get('Integer1'), self.style14)
            else:
                self.ws.write(self.rowsize, 39, self.update_custom_details_dict.get('Integer1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer2[loop] == self.update_custom_details_dict.get('Integer2'):
                if self.xl_update_Integer2[loop] is None:
                    self.ws.write(self.rowsize, 40, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 40, self.update_custom_details_dict.get('Integer2'), self.style14)
            else:
                self.ws.write(self.rowsize, 40, self.update_custom_details_dict.get('Integer2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer3[loop] == self.update_custom_details_dict.get('Integer3'):
                if self.xl_update_Integer3[loop] is None:
                    self.ws.write(self.rowsize, 41, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 41, self.update_custom_details_dict.get('Integer3'), self.style14)
            else:
                self.ws.write(self.rowsize, 41, self.update_custom_details_dict.get('Integer3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer4[loop] == self.update_custom_details_dict.get('Integer4'):
                if self.xl_update_Integer4[loop] is None:
                    self.ws.write(self.rowsize, 42, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 42, self.update_custom_details_dict.get('Integer4'), self.style14)
            else:
                self.ws.write(self.rowsize, 42, self.update_custom_details_dict.get('Integer4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer5[loop] == self.update_custom_details_dict.get('Integer5'):
                if self.xl_update_Integer5[loop] is None:
                    self.ws.write(self.rowsize, 43, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 43, self.update_custom_details_dict.get('Integer5'), self.style14)
            else:
                self.ws.write(self.rowsize, 43, self.update_custom_details_dict.get('Integer5'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer6[loop] == self.update_custom_details_dict.get('Integer6'):
                if self.xl_update_Integer6[loop] is None:
                    self.ws.write(self.rowsize, 44, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 44, self.update_custom_details_dict.get('Integer6'), self.style14)
            else:
                self.ws.write(self.rowsize, 44, self.update_custom_details_dict.get('Integer6'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer7[loop] == self.update_custom_details_dict.get('Integer7'):
                if self.xl_update_Integer7[loop] is None:
                    self.ws.write(self.rowsize, 45, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 45, self.update_custom_details_dict.get('Integer7'), self.style14)
            else:
                self.ws.write(self.rowsize, 45, self.update_custom_details_dict.get('Integer7'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer8[loop] == self.update_custom_details_dict.get('Integer8'):
                if self.xl_update_Integer8[loop] is None:
                    self.ws.write(self.rowsize, 46, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 46, self.update_custom_details_dict.get('Integer8'), self.style14)
            else:
                self.ws.write(self.rowsize, 46, self.update_custom_details_dict.get('Integer8'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer9[loop] == self.update_custom_details_dict.get('Integer9'):
                if self.xl_update_Integer9[loop] is None:
                    self.ws.write(self.rowsize, 47, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 47, self.update_custom_details_dict.get('Integer9'), self.style14)
            else:
                self.ws.write(self.rowsize, 47, self.update_custom_details_dict.get('Integer9'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer10[loop] == self.update_custom_details_dict.get('Integer10'):
                if self.xl_update_Integer10[loop] is None:
                    self.ws.write(self.rowsize, 48, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 48, self.update_custom_details_dict.get('Integer10'), self.style14)
            else:
                self.ws.write(self.rowsize, 48, self.update_custom_details_dict.get('Integer10'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer11[loop] == self.update_custom_details_dict.get('Integer11'):
                if self.xl_update_Integer11[loop] is None:
                    self.ws.write(self.rowsize, 49, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 49, self.update_custom_details_dict.get('Integer11'), self.style14)
            else:
                self.ws.write(self.rowsize, 49, self.update_custom_details_dict.get('Integer11'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer12[loop] == self.update_custom_details_dict.get('Integer12'):
                if self.xl_update_Integer12[loop] is None:
                    self.ws.write(self.rowsize, 50, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 50, self.update_custom_details_dict.get('Integer12'), self.style14)
            else:
                self.ws.write(self.rowsize, 50, self.update_custom_details_dict.get('Integer12'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer13[loop] == self.update_custom_details_dict.get('Integer13'):
                if self.xl_update_Integer13[loop] is None:
                    self.ws.write(self.rowsize, 51, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 51, self.update_custom_details_dict.get('Integer13'), self.style14)
            else:
                self.ws.write(self.rowsize, 51, self.update_custom_details_dict.get('Integer13'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer14[loop] == self.update_custom_details_dict.get('Integer14'):
                if self.xl_update_Integer14[loop] is None:
                    self.ws.write(self.rowsize, 52, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 52, self.update_custom_details_dict.get('Integer14'), self.style14)
            else:
                self.ws.write(self.rowsize, 52, self.update_custom_details_dict.get('Integer14'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Integer15[loop] == self.update_custom_details_dict.get('Integer15'):
                if self.xl_update_Integer15[loop] is None:
                    self.ws.write(self.rowsize, 53, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 53, self.update_custom_details_dict.get('Integer15'), self.style14)
            else:
                self.ws.write(self.rowsize, 53, self.update_custom_details_dict.get('Integer15'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text1[loop] == self.update_custom_details_dict.get('Text1'):
                if self.xl_update_Text1[loop] is None:
                    self.ws.write(self.rowsize, 54, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 54, self.update_custom_details_dict.get('Text1'), self.style14)
            else:
                self.ws.write(self.rowsize, 54, self.update_custom_details_dict.get('Text1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text2[loop] == self.update_custom_details_dict.get('Text2'):
                if self.xl_update_Text2[loop] is None:
                    self.ws.write(self.rowsize, 55, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 55, self.update_custom_details_dict.get('Text2'), self.style14)
            else:
                self.ws.write(self.rowsize, 55, self.update_custom_details_dict.get('Text2'), self.style3)

        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text3[loop] == self.update_custom_details_dict.get('Text3'):
                if self.xl_update_Text3[loop] is None:
                    self.ws.write(self.rowsize, 56, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 56, self.update_custom_details_dict.get('Text3'), self.style14)
            else:
                self.ws.write(self.rowsize, 56, self.update_custom_details_dict.get('Text3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text4[loop] == self.update_custom_details_dict.get('Text4'):
                if self.xl_update_Text4[loop] is None:
                    self.ws.write(self.rowsize, 57, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 57, self.update_custom_details_dict.get('Text4'), self.style14)
            else:
                self.ws.write(self.rowsize, 57, self.update_custom_details_dict.get('Text4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text5[loop] == self.update_custom_details_dict.get('Text5'):
                if self.xl_update_Text5[loop] is None:
                    self.ws.write(self.rowsize, 58, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 58, self.update_custom_details_dict.get('Text5'), self.style14)
            else:
                self.ws.write(self.rowsize, 58, self.update_custom_details_dict.get('Text5'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text6[loop] == self.update_custom_details_dict.get('Text6'):
                if self.xl_update_Text6[loop] is None:
                    self.ws.write(self.rowsize, 59, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 59, self.update_custom_details_dict.get('Text6'), self.style14)
            else:
                self.ws.write(self.rowsize, 59, self.update_custom_details_dict.get('Text6'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text7[loop] == self.update_custom_details_dict.get('Text7'):
                if self.xl_update_Text7[loop] is None:
                    self.ws.write(self.rowsize, 60, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 60, self.update_custom_details_dict.get('Text7'), self.style14)
            else:
                self.ws.write(self.rowsize, 60, self.update_custom_details_dict.get('Text7'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text8[loop] == self.update_custom_details_dict.get('Text8'):
                if self.xl_update_Text8[loop] is None:
                    self.ws.write(self.rowsize, 61, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 61, self.update_custom_details_dict.get('Text8'), self.style14)
            else:
                self.ws.write(self.rowsize, 61, self.update_custom_details_dict.get('Text8'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text9[loop] == self.update_custom_details_dict.get('Text9'):
                if self.xl_update_Text9[loop] is None:
                    self.ws.write(self.rowsize, 62, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 62, self.update_custom_details_dict.get('Text9'), self.style14)
            else:
                self.ws.write(self.rowsize, 62, self.update_custom_details_dict.get('Text9'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text10[loop] == self.update_custom_details_dict.get('Text10'):
                if self.xl_update_Text10[loop] is None:
                    self.ws.write(self.rowsize, 63, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 63, self.update_custom_details_dict.get('Text10'), self.style14)
            else:
                self.ws.write(self.rowsize, 63, self.update_custom_details_dict.get('Text10'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text11[loop] == self.update_custom_details_dict.get('Text11'):
                if self.xl_update_Text11[loop] is None:
                    self.ws.write(self.rowsize, 64, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 64, self.update_custom_details_dict.get('Text11'), self.style14)
            else:
                self.ws.write(self.rowsize, 64, self.update_custom_details_dict.get('Text11'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text12[loop] == self.update_custom_details_dict.get('Text12'):
                if self.xl_update_Text12[loop] is None:
                    self.ws.write(self.rowsize, 65, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 65, self.update_custom_details_dict.get('Text12'), self.style14)
            else:
                self.ws.write(self.rowsize, 65, self.update_custom_details_dict.get('Text12'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text13[loop] == self.update_custom_details_dict.get('Text13'):
                if self.xl_update_Text13[loop] is None:
                    self.ws.write(self.rowsize, 66, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 66, self.update_custom_details_dict.get('Text13'), self.style14)
            else:
                self.ws.write(self.rowsize, 66, self.update_custom_details_dict.get('Text13'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text14[loop] == self.update_custom_details_dict.get('Text14'):
                if self.xl_update_Text14[loop] is None:
                    self.ws.write(self.rowsize, 67, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 67, self.update_custom_details_dict.get('Text14'), self.style14)
            else:
                self.ws.write(self.rowsize, 67, self.update_custom_details_dict.get('Text14'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_Text15[loop] == self.update_custom_details_dict.get('Text15'):
                if self.xl_update_Text15[loop] is None:
                    self.ws.write(self.rowsize, 68, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 68, self.update_custom_details_dict.get('Text15'), self.style14)
            else:
                self.ws.write(self.rowsize, 68, self.update_custom_details_dict.get('Text15'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TextArea1[loop] == self.update_custom_details_dict.get('TextArea1'):
                if self.xl_update_TextArea1[loop] is None:
                    self.ws.write(self.rowsize, 69, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 69, self.update_custom_details_dict.get('TextArea1'), self.style14)
            else:
                self.ws.write(self.rowsize, 69, self.update_custom_details_dict.get('TextArea1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TextArea2[loop] == self.update_custom_details_dict.get('TextArea2'):
                if self.xl_update_TextArea2[loop] is None:
                    self.ws.write(self.rowsize, 70, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 70, self.update_custom_details_dict.get('TextArea2'), self.style14)
            else:
                self.ws.write(self.rowsize, 70, self.update_custom_details_dict.get('TextArea2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TextArea3[loop] == self.update_custom_details_dict.get('TextArea3'):
                if self.xl_update_TextArea3[loop] is None:
                    self.ws.write(self.rowsize, 71, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 71, self.update_custom_details_dict.get('TextArea3'), self.style14)
            else:
                self.ws.write(self.rowsize, 71, self.update_custom_details_dict.get('TextArea3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TextArea4[loop] == self.update_custom_details_dict.get('TextArea4'):
                if self.xl_update_TextArea4[loop] is None:
                    self.ws.write(self.rowsize, 72, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 72, self.update_custom_details_dict.get('TextArea4'), self.style14)
            else:
                self.ws.write(self.rowsize, 72, self.update_custom_details_dict.get('TextArea4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DateCustomField1[loop] == self.update_custom_details_dict.get('DateCustomField1'):
                if self.xl_update_DateCustomField1[loop] is None:
                    self.ws.write(self.rowsize, 73, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 73,
                                  self.update_custom_details_dict.get('DateCustomField1'), self.style14)
            elif self.update_custom_details_dict.get('DateCustomField1'):
                if '00:00:00' in self.update_custom_details_dict.get('DateCustomField1'):
                    self.ws.write(self.rowsize, 73,
                                  self.update_custom_details_dict.get('DateCustomField1'), self.style7)
            else:
                self.ws.write(self.rowsize, 73, self.update_custom_details_dict.get('DateCustomField1'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DateCustomField2[loop] == self.update_custom_details_dict.get('DateCustomField2'):
                if self.xl_update_DateCustomField2[loop] is None:
                    self.ws.write(self.rowsize, 74, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 74,
                                  self.update_custom_details_dict.get('DateCustomField2'), self.style14)
            elif self.update_custom_details_dict.get('DateCustomField2'):
                if '00:00:00' in self.update_custom_details_dict.get('DateCustomField2'):
                    self.ws.write(self.rowsize, 74,
                                  self.update_custom_details_dict.get('DateCustomField2'), self.style7)
            else:
                self.ws.write(self.rowsize, 74, self.update_custom_details_dict.get('DateCustomField2'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DateCustomField3[loop] == self.update_custom_details_dict.get('DateCustomField3'):
                if self.xl_update_DateCustomField3[loop] is None:
                    self.ws.write(self.rowsize, 75, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 75,
                                  self.update_custom_details_dict.get('DateCustomField3'), self.style14)
            elif self.update_custom_details_dict.get('DateCustomField3'):
                if '00:00:00' in self.update_custom_details_dict.get('DateCustomField3'):
                    self.ws.write(self.rowsize, 75,
                                  self.update_custom_details_dict.get('DateCustomField3'), self.style7)
            else:
                self.ws.write(self.rowsize, 75, self.update_custom_details_dict.get('DateCustomField3'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DateCustomField4[loop] == self.update_custom_details_dict.get('DateCustomField4'):
                if self.error_description:
                    self.ws.write(self.rowsize, 76, self.error_description, self.style3)
                elif self.xl_update_DateCustomField4[loop] is None:
                    self.ws.write(self.rowsize, 76, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 76,
                                  self.update_custom_details_dict.get('DateCustomField4'), self.style14)
            elif self.update_custom_details_dict.get('DateCustomField4'):
                if '00:00:00' in self.update_custom_details_dict.get('DateCustomField4'):
                    self.ws.write(self.rowsize, 76,
                                  self.update_custom_details_dict.get('DateCustomField4'), self.style7)
            else:
                self.ws.write(self.rowsize, 76, self.update_custom_details_dict.get('DateCustomField4'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_DateCustomField5[loop] == self.update_custom_details_dict.get('DateCustomField5'):
                if self.xl_update_DateCustomField5[loop] is None:
                    self.ws.write(self.rowsize, 77, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 77,
                                  self.update_custom_details_dict.get('DateCustomField5'), self.style14)
            elif self.update_custom_details_dict.get('DateCustomField5'):
                if '00:00:00' in self.update_custom_details_dict.get('DateCustomField5'):
                    self.ws.write(self.rowsize, 77,
                                  self.update_custom_details_dict.get('DateCustomField5'), self.style7)
            else:
                self.ws.write(self.rowsize, 77, self.update_custom_details_dict.get('DateCustomField5'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        # --------------------------------------------------------------------------------------------------------------
        if self.update_custom_details_dict.get('TrueFalse1'):
            truefalse1 = 'true'
        else:
            truefalse1 = 'false'
        if self.update_custom_details_dict.get('TrueFalse2'):
            truefalse2 = 'true'
        else:
            truefalse2 = 'false'
        if self.update_custom_details_dict.get('TrueFalse3'):
            truefalse3 = 'true'
        else:
            truefalse3 = 'false'
        if self.update_custom_details_dict.get('TrueFalse4'):
            truefalse4 = 'true'
        else:
            truefalse4 = 'false'
        if self.update_custom_details_dict.get('TrueFalse5'):
            truefalse5 = 'true'
        else:
            truefalse5 = 'false'
        # --------------------------------------------------------------------------------------------------------------
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TrueFalse1[loop] == truefalse1:
                if self.xl_update_TrueFalse1[loop] is None:
                    self.ws.write(self.rowsize, 78, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 78, truefalse1, self.style14)
            else:
                self.ws.write(self.rowsize, 78, truefalse1, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TrueFalse2[loop] == truefalse2:
                if self.xl_update_TrueFalse2[loop] is None:
                    self.ws.write(self.rowsize, 79, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 79, truefalse2, self.style14)
            else:
                self.ws.write(self.rowsize, 79, truefalse2, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TrueFalse3[loop] == truefalse3:
                if self.xl_update_TrueFalse3[loop] is None:
                    self.ws.write(self.rowsize, 80, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 80, truefalse3, self.style14)
            else:
                self.ws.write(self.rowsize, 80, truefalse3, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TrueFalse4[loop] == truefalse4:
                if self.xl_update_TrueFalse4[loop] is None:
                    self.ws.write(self.rowsize, 81, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 81, truefalse4, self.style14)
            else:
                self.ws.write(self.rowsize, 81, truefalse4, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if not self.error_description:
            if self.xl_update_TrueFalse5[loop] == truefalse5:
                if self.xl_update_TrueFalse5[loop] is None:
                    self.ws.write(self.rowsize, 82, 'Empty', self.style14)
                else:
                    self.ws.write(self.rowsize, 82, truefalse5, self.style14)
            else:
                self.ws.write(self.rowsize, 82, truefalse5, self.style3)
        # --------------------------------------------------------------------------------------------------------------
        # if not self.error_description:
        #     if self.xl_update_TotalExperienceInYears[loop] == self.update_personal_details_dict \
        #             .get('TotalExperienceInYears'):
        #         if self.xl_update_TotalExperienceInYears[loop] is None:
        #             self.ws.write(self.rowsize, 83, 'Empty', self.style14)
        #         else:
        #             self.ws.write(self.rowsize, 83,
        #                           self.update_personal_details_dict.get('TotalExperienceInYears'), self.style14)
        #     else:
        #         self.ws.write(self.rowsize, 83,
        #                       self.update_personal_details_dict.get('TotalExperienceInYears'), self.style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.error_description:
            if self.xl_updated_expected_message[loop] == self.error_description:
                self.ws.write(self.rowsize, 83, self.error_description, self.style14)
            else:
                self.ws.write(self.rowsize, 83, self.error_description, self.style3)

        # --------------------- overall success cases ------------------------------------
        if self.success_case_01 == 'Pass':
            self.Actual_Success_case.append(self.success_case_01)
        if self.success_case_02 == 'Pass':
            self.Actual_Success_case.append(self.success_case_02)

        self.rowsize += 1
        Object.wb_Result.save(output_paths.outputpaths['Update_Candidate_Output_sheet'])

    def overall_status(self):
        self.ws.write(self.final_status_rowsize, 0, 'Update Candidates', self.style23)
        if self.Expected_success_cases == self.Actual_Success_case:
            self.ws.write(self.final_status_rowsize, 1, 'Pass', self.style24)
        else:
            self.ws.write(self.final_status_rowsize, 1, 'Fail', self.style25)
        self.ws.write(self.final_status_rowsize, 2, 'Start Time', self.style23)
        self.ws.write(self.final_status_rowsize, 3, self.start_time, self.style26)
        self.ws.write(0, 4, 'Lambda', self.style23)
        self.ws.write(0, 5, self.calling_lambda, self.style24)
        Object.wb_Result.save(output_paths.outputpaths['Update_Candidate_Output_sheet'])


Object = UpdateCandidate()
Object.excel_headers()
Object.read_excel()
Total_count = len(Object.xl_update_candidate_id)
print("Total count ::", Total_count)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.update_candidate(looping)
        Object.mobile_update(looping)
        Object.candidate_get_by_id_details(looping)
        Object.out_file(looping)

        # ------------------
        # Making Dict Empty
        # ------------------
        Object.update_personal_details_dict = {}
        Object.update_source_details_dict = {}
        Object.update_custom_details_dict = {}
        Object.update_social_details_dict = {}
        Object.update_primary_skills_dict = {}
        Object.update_secondary_skills_dict = {}
        Object.update_candidate_preference_dict = {}
        Object.api_updated_CID = {}
        Object.error_description = {}
        Object.success_case_01 = {}
        Object.success_case_02 = {}
        Object.headers = {}

Object.overall_status()
