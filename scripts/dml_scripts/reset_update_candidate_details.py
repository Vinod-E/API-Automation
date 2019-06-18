import xlrd
import json
import requests
from hpro_automation import (api, login, input_paths)


class ResetCandidateDetails(login.CRPOLogin):

    def __init__(self):

        super(ResetCandidateDetails, self).__init__()

        self.xl_reset_candidate_id = []
        self.xl_reset_Name = []
        self.xl_reset_FirstName = []
        self.xl_reset_MiddleName = []
        self.xl_reset_LastName = []
        self.xl_reset_TotalExperienceInMonths = []
        self.xl_reset_passport = []
        self.xl_reset_CurrentCTC = []
        self.xl_reset_Mobile1 = []
        self.xl_reset_PanNo = []
        self.xl_reset_Martial = []
        self.xl_reset_DOB = []
        self.xl_reset_CurrentLocId = []
        self.xl_reset_CurrentLocText = []
        self.xl_reset_Address1 = []
        self.xl_reset_ExpertiseId1 = []
        self.xl_reset_PhoneOffice = []
        self.xl_reset_Gender = []
        self.xl_reset_TotalExperienceInYears = []
        self.xl_reset_Email2 = []
        self.xl_reset_Email1 = []
        self.xl_reset_CurrencyType = []
        self.xl_reset_Sensitivity = []
        self.xl_reset_Nationality = []
        self.xl_reset_Country = []
        self.xl_reset_StatusId = []
        self.xl_reset_USN = []
        self.xl_reset_AadharNo = []
        self.xl_reset_HierarchyId = []
        self.xl_reset_DesiredSalaryFrom = []
        self.xl_reset_DesiredSalaryTo = []
        self.xl_reset_NoticePeriod = []
        self.xl_reset_SourceId = []
        self.xl_reset_SourceType = []
        self.xl_reset_Skill1 = []
        self.xl_reset_Skill2 = []
        self.xl_reset_LinkedInLink = []
        self.xl_reset_FacebookLink = []
        self.xl_reset_TwitterLink = []
        self.xl_reset_Integer1 = []
        self.xl_reset_Integer2 = []
        self.xl_reset_Integer3 = []
        self.xl_reset_Integer4 = []
        self.xl_reset_Integer5 = []
        self.xl_reset_Integer6 = []
        self.xl_reset_Integer7 = []
        self.xl_reset_Integer8 = []
        self.xl_reset_Integer9 = []
        self.xl_reset_Integer10 = []
        self.xl_reset_Integer11 = []
        self.xl_reset_Integer12 = []
        self.xl_reset_Integer13 = []
        self.xl_reset_Integer14 = []
        self.xl_reset_Integer15 = []
        self.xl_reset_Text1 = []
        self.xl_reset_Text2 = []
        self.xl_reset_Text3 = []
        self.xl_reset_Text4 = []
        self.xl_reset_Text5 = []
        self.xl_reset_Text6 = []
        self.xl_reset_Text7 = []
        self.xl_reset_Text8 = []
        self.xl_reset_Text9 = []
        self.xl_reset_Text10 = []
        self.xl_reset_Text11 = []
        self.xl_reset_Text12 = []
        self.xl_reset_Text13 = []
        self.xl_reset_Text14 = []
        self.xl_reset_Text15 = []
        self.xl_reset_TextArea1 = []
        self.xl_reset_TextArea2 = []
        self.xl_reset_TextArea3 = []
        self.xl_reset_TextArea4 = []
        self.xl_reset_TrueFalse1 = []
        self.xl_reset_TrueFalse2 = []
        self.xl_reset_TrueFalse3 = []
        self.xl_reset_TrueFalse4 = []
        self.xl_reset_TrueFalse5 = []
        self.xl_reset_DateCustomField1 = []
        self.xl_reset_DateCustomField2 = []
        self.xl_reset_DateCustomField3 = []
        self.xl_reset_DateCustomField4 = []
        self.xl_reset_DateCustomField5 = []

        self.headers = {}

    def read_excel(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['reset_candi_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if rows[0]:
                self.xl_reset_candidate_id.append(int(rows[0]))

            if rows[1]:
                self.xl_reset_Name.append(rows[1])

            if not rows[2]:
                self.xl_reset_FirstName.append(None)
            else:
                self.xl_reset_FirstName.append(rows[2])

            if not rows[3]:
                self.xl_reset_MiddleName.append(None)
            else:
                self.xl_reset_MiddleName.append(rows[3])

            if not rows[4]:
                self.xl_reset_LastName.append(None)
            else:
                self.xl_reset_LastName.append(rows[4])

            if not rows[5]:
                self.xl_reset_TotalExperienceInMonths.append(None)
            else:
                self.xl_reset_TotalExperienceInMonths.append(int(rows[5]))

            if not rows[6]:
                self.xl_reset_passport.append(None)
            else:
                self.xl_reset_passport.append(str(rows[6]))

            if not rows[7]:
                self.xl_reset_CurrentCTC.append(None)
            else:
                self.xl_reset_CurrentCTC.append(int(rows[7]))

            if not rows[8]:
                self.xl_reset_Mobile1.append(None)
            else:
                self.xl_reset_Mobile1.append(str(int(rows[8])))

            if not rows[9]:
                self.xl_reset_PanNo.append(None)
            else:
                self.xl_reset_PanNo.append(str(rows[9]))

            if not rows[10]:
                self.xl_reset_Martial.append(None)
            else:
                self.xl_reset_Martial.append(int(rows[10]))

            if not rows[11]:
                self.xl_reset_DOB.append(None)
            else:
                self.xl_reset_DOB.append(str(rows[11]))

            if not rows[12]:
                self.xl_reset_CurrentLocId.append(None)
            else:
                self.xl_reset_CurrentLocId.append(int(rows[12]))

            if not rows[13]:
                self.xl_reset_CurrentLocText.append(None)
            else:
                self.xl_reset_CurrentLocText.append(str(rows[13]))

            if not rows[14]:
                self.xl_reset_Address1.append(None)
            else:
                self.xl_reset_Address1.append(str(rows[14]))

            if not rows[15]:
                self.xl_reset_ExpertiseId1.append(None)
            else:
                self.xl_reset_ExpertiseId1.append(int(rows[15]))

            if not rows[16]:
                self.xl_reset_PhoneOffice.append(None)
            else:
                self.xl_reset_PhoneOffice.append(str(int(rows[16])))

            if not rows[17]:
                self.xl_reset_Gender.append(None)
            else:
                self.xl_reset_Gender.append(int(rows[17]))

            if not rows[18]:
                self.xl_reset_TotalExperienceInYears.append(None)
            else:
                self.xl_reset_TotalExperienceInYears.append(int(rows[18]))

            if not rows[19]:
                self.xl_reset_Email2.append(None)
            else:
                self.xl_reset_Email2.append(str(rows[19]))

            if not rows[20]:
                self.xl_reset_Email1.append(None)
            else:
                self.xl_reset_Email1.append(str(rows[20]))

            if not rows[21]:
                self.xl_reset_CurrencyType.append(None)
            else:
                self.xl_reset_CurrencyType.append(str(rows[21]))

            if not rows[22]:
                self.xl_reset_Sensitivity.append(None)
            else:
                self.xl_reset_Sensitivity.append(int(rows[22]))

            if not rows[23]:
                self.xl_reset_Nationality.append(None)
            else:
                self.xl_reset_Nationality.append(int(rows[23]))

            if not rows[24]:
                self.xl_reset_Country.append(None)
            else:
                self.xl_reset_Country.append(int(rows[24]))

            if not rows[25]:
                self.xl_reset_StatusId.append(None)
            else:
                self.xl_reset_StatusId.append(int(rows[25]))

            if not rows[26]:
                self.xl_reset_USN.append(None)
            else:
                self.xl_reset_USN.append(str(rows[26]))

            if not rows[27]:
                self.xl_reset_AadharNo.append(None)
            else:
                self.xl_reset_AadharNo.append(str(rows[27]))

            if not rows[28]:
                self.xl_reset_HierarchyId.append(None)
            else:
                self.xl_reset_HierarchyId.append(int(rows[28]))

            if not rows[29]:
                self.xl_reset_DesiredSalaryFrom.append(None)
            else:
                self.xl_reset_DesiredSalaryFrom.append(int(rows[29]))

            if not rows[30]:
                self.xl_reset_DesiredSalaryTo.append(None)
            else:
                self.xl_reset_DesiredSalaryTo.append(int(rows[30]))

            if not rows[31]:
                self.xl_reset_NoticePeriod.append(None)
            else:
                self.xl_reset_NoticePeriod.append(int(rows[31]))

            if not rows[32]:
                self.xl_reset_SourceId.append(None)
            else:
                self.xl_reset_SourceId.append(int(rows[32]))

            if not rows[33]:
                self.xl_reset_SourceType.append(None)
            else:
                self.xl_reset_SourceType.append(int(rows[33]))

            if not rows[34]:
                self.xl_reset_Skill1.append(None)
            else:
                self.xl_reset_Skill1.append(int(rows[34]))

            if not rows[35]:
                self.xl_reset_Skill2.append(None)
            else:
                self.xl_reset_Skill2.append(int(rows[35]))

            if not rows[36]:
                self.xl_reset_LinkedInLink.append(None)
            else:
                self.xl_reset_LinkedInLink.append(str(rows[36]))

            if not rows[37]:
                self.xl_reset_FacebookLink.append(None)
            else:
                self.xl_reset_FacebookLink.append(str(rows[37]))

            if not rows[38]:
                self.xl_reset_TwitterLink.append(None)
            else:
                self.xl_reset_TwitterLink.append(str(rows[38]))

            if not rows[39]:
                self.xl_reset_Integer1.append(None)
            else:
                self.xl_reset_Integer1.append(int(rows[39]))

            if not rows[40]:
                self.xl_reset_Integer2.append(None)
            else:
                self.xl_reset_Integer2.append(int(rows[40]))

            if not rows[41]:
                self.xl_reset_Integer3.append(None)
            else:
                self.xl_reset_Integer3.append(int(rows[41]))

            if not rows[42]:
                self.xl_reset_Integer4.append(None)
            else:
                self.xl_reset_Integer4.append(int(rows[42]))

            if not rows[43]:
                self.xl_reset_Integer5.append(None)
            else:
                self.xl_reset_Integer5.append(int(rows[43]))

            if not rows[44]:
                self.xl_reset_Integer6.append(None)
            else:
                self.xl_reset_Integer6.append(int(rows[44]))

            if not rows[45]:
                self.xl_reset_Integer7.append(None)
            else:
                self.xl_reset_Integer7.append(int(rows[45]))

            if not rows[46]:
                self.xl_reset_Integer8.append(None)
            else:
                self.xl_reset_Integer8.append(int(rows[46]))

            if not rows[47]:
                self.xl_reset_Integer9.append(None)
            else:
                self.xl_reset_Integer9.append(int(rows[47]))

            if not rows[48]:
                self.xl_reset_Integer10.append(None)
            else:
                self.xl_reset_Integer10.append(int(rows[48]))

            if not rows[49]:
                self.xl_reset_Integer11.append(None)
            else:
                self.xl_reset_Integer11.append(int(rows[49]))

            if not rows[50]:
                self.xl_reset_Integer12.append(None)
            else:
                self.xl_reset_Integer12.append(int(rows[50]))

            if not rows[51]:
                self.xl_reset_Integer13.append(None)
            else:
                self.xl_reset_Integer13.append(int(rows[51]))

            if not rows[52]:
                self.xl_reset_Integer14.append(None)
            else:
                self.xl_reset_Integer14.append(int(rows[52]))

            if not rows[53]:
                self.xl_reset_Integer15.append(None)
            else:
                self.xl_reset_Integer15.append(int(rows[53]))

            if not rows[54]:
                self.xl_reset_Text1.append(None)
            else:
                self.xl_reset_Text1.append(rows[54])

            if not rows[55]:
                self.xl_reset_Text2.append(None)
            else:
                self.xl_reset_Text2.append(rows[55])

            if not rows[56]:
                self.xl_reset_Text3.append(None)
            else:
                self.xl_reset_Text3.append(rows[56])

            if not rows[57]:
                self.xl_reset_Text4.append(None)
            else:
                self.xl_reset_Text4.append(rows[57])

            if not rows[58]:
                self.xl_reset_Text5.append(None)
            else:
                self.xl_reset_Text5.append(rows[58])

            if not rows[59]:
                self.xl_reset_Text6.append(None)
            else:
                self.xl_reset_Text6.append(rows[59])

            if not rows[60]:
                self.xl_reset_Text7.append(None)
            else:
                self.xl_reset_Text7.append(rows[60])

            if not rows[61]:
                self.xl_reset_Text8.append(None)
            else:
                self.xl_reset_Text8.append(rows[61])

            if not rows[62]:
                self.xl_reset_Text9.append(None)
            else:
                self.xl_reset_Text9.append(rows[62])

            if not rows[63]:
                self.xl_reset_Text10.append(None)
            else:
                self.xl_reset_Text10.append(rows[63])

            if not rows[64]:
                self.xl_reset_Text11.append(None)
            else:
                self.xl_reset_Text11.append(rows[64])

            if not rows[65]:
                self.xl_reset_Text12.append(None)
            else:
                self.xl_reset_Text12.append(rows[65])

            if not rows[66]:
                self.xl_reset_Text13.append(None)
            else:
                self.xl_reset_Text13.append(rows[66])

            if not rows[67]:
                self.xl_reset_Text14.append(None)
            else:
                self.xl_reset_Text14.append(rows[67])

            if not rows[68]:
                self.xl_reset_Text15.append(None)
            else:
                self.xl_reset_Text15.append(rows[68])

            if not rows[69]:
                self.xl_reset_TextArea1.append(None)
            else:
                self.xl_reset_TextArea1.append(rows[69])

            if not rows[70]:
                self.xl_reset_TextArea2.append(None)
            else:
                self.xl_reset_TextArea2.append(rows[70])

            if not rows[71]:
                self.xl_reset_TextArea3.append(None)
            else:
                self.xl_reset_TextArea3.append(rows[71])

            if not rows[72]:
                self.xl_reset_TextArea4.append(None)
            else:
                self.xl_reset_TextArea4.append(rows[72])

            if not rows[72]:
                self.xl_reset_TextArea4.append(None)
            else:
                self.xl_reset_TextArea4.append(rows[72])

            if not rows[73]:
                self.xl_reset_DateCustomField1.append(None)
            else:
                self.xl_reset_DateCustomField1.append(rows[73])

            if not rows[74]:
                self.xl_reset_DateCustomField2.append(None)
            else:
                self.xl_reset_DateCustomField2.append(rows[74])

            if not rows[75]:
                self.xl_reset_DateCustomField3.append(None)
            else:
                self.xl_reset_DateCustomField3.append(rows[75])

            if not rows[76]:
                self.xl_reset_DateCustomField4.append(None)
            else:
                self.xl_reset_DateCustomField4.append(rows[76])

            if not rows[77]:
                self.xl_reset_DateCustomField5.append(None)
            else:
                self.xl_reset_DateCustomField5.append(rows[77])

            if not rows[78]:
                self.xl_reset_TrueFalse1.append(None)
            else:
                self.xl_reset_TrueFalse1.append(rows[78])

            if not rows[79]:
                self.xl_reset_TrueFalse2.append(None)
            else:
                self.xl_reset_TrueFalse2.append(rows[79])

            if not rows[80]:
                self.xl_reset_TrueFalse3.append(None)
            else:
                self.xl_reset_TrueFalse3.append(rows[80])

            if not rows[81]:
                self.xl_reset_TrueFalse4.append(None)
            else:
                self.xl_reset_TrueFalse4.append(rows[81])

            if not rows[82]:
                self.xl_reset_TrueFalse5.append(None)
            else:
                self.xl_reset_TrueFalse5.append(rows[82])

    def reset_candidate_details(self, loop):

        self.lambda_function('update_candidate_details')
        self.headers['APP-NAME'] = 'crpo'

        if self.xl_reset_TrueFalse1[loop] == 'true':
            truefalse1 = True
        else:
            truefalse1 = False
        if self.xl_reset_TrueFalse2[loop] == 'true':
            truefalse2 = True
        else:
            truefalse2 = False
        if self.xl_reset_TrueFalse3[loop] == 'true':
            truefalse3 = True
        else:
            truefalse3 = False
        if self.xl_reset_TrueFalse4[loop] == 'true':
            truefalse4 = True
        else:
            truefalse4 = False
        if self.xl_reset_TrueFalse5[loop] == 'true':
            truefalse5 = True
        else:
            truefalse5 = False

        request = {
            "PersonalDetails": {
                "Name": self.xl_reset_Name[loop],
                "FirstName": self.xl_reset_FirstName[loop],
                "MiddleName": self.xl_reset_MiddleName[loop],
                "LastName": self.xl_reset_LastName[loop],
                "TotalExperienceInMonths": self.xl_reset_TotalExperienceInMonths[loop],
                "PassportNo": self.xl_reset_passport[loop],
                "CurrentCTC": self.xl_reset_CurrentCTC[loop],
                "PanNo": self.xl_reset_PanNo[loop],
                "MaritalStatus": self.xl_reset_Martial[loop],
                "DateOfBirth": self.xl_reset_DOB[loop],
                "CurrentLocationId": self.xl_reset_CurrentLocId[loop],
                "CurrentLocationText": self.xl_reset_CurrentLocText[loop],
                "Address1": self.xl_reset_Address1[loop],
                "ExpertiseId1": self.xl_reset_ExpertiseId1[loop],
                "PhoneOffice": self.xl_reset_PhoneOffice[loop],
                "Gender": self.xl_reset_Gender[loop],
                "TotalExperienceInYears": self.xl_reset_TotalExperienceInYears[loop],
                "Email2": self.xl_reset_Email2[loop],
                "Email1": self.xl_reset_Email1[loop],
                "CurrencyType": self.xl_reset_CurrencyType[loop],
                "Sensitivity": self.xl_reset_Sensitivity[loop],
                "Nationality": self.xl_reset_Nationality[loop],
                "Country": self.xl_reset_Country[loop],
                "StatusId": self.xl_reset_StatusId[loop],
                "USN": self.xl_reset_USN[loop],
                "AadhaarNo": self.xl_reset_AadharNo[loop],
                "HierarchyId": self.xl_reset_HierarchyId[loop]
            },
            "PreferenceDetails": {"CurrencyType": self.xl_reset_CurrencyType[loop],
                                  "DesiredSalaryFrom": self.xl_reset_DesiredSalaryFrom[loop],
                                  "DesiredSalaryTo": self.xl_reset_DesiredSalaryTo[loop],
                                  "NoticePeriod": self.xl_reset_NoticePeriod[loop]},
            "SourceDetails": {
                "SourceId": self.xl_reset_SourceId[loop]
            },
            "SocialDetails": {
                "LinkedInLink": self.xl_reset_LinkedInLink[loop],
                "FacebookLink": self.xl_reset_FacebookLink[loop],
                "TwitterLink": self.xl_reset_TwitterLink[loop]
            },
            "SkillsDetails": {
                "KeySkills": [{
                    "Id": self.xl_reset_Skill1[loop]
                }],
                "SecondaryKeySkills": [{
                    "Id": self.xl_reset_Skill2[loop]
                }]
            },
            "CustomDetails": {
                "Integer1": self.xl_reset_Integer1[loop],
                "Integer2": self.xl_reset_Integer2[loop],
                "Integer3": self.xl_reset_Integer3[loop],
                "Integer4": self.xl_reset_Integer4[loop],
                "Integer5": self.xl_reset_Integer5[loop],
                "Integer6": self.xl_reset_Integer6[loop],
                "Integer7": self.xl_reset_Integer7[loop],
                "Integer8": self.xl_reset_Integer8[loop],
                "Integer9": self.xl_reset_Integer9[loop],
                "Integer10": self.xl_reset_Integer10[loop],
                "Integer11": self.xl_reset_Integer11[loop],
                "Integer12": self.xl_reset_Integer12[loop],
                "Integer13": self.xl_reset_Integer13[loop],
                "Integer14": self.xl_reset_Integer14[loop],
                "Integer15": self.xl_reset_Integer15[loop],
                "Text1": self.xl_reset_Text1[loop],
                "Text2": self.xl_reset_Text2[loop],
                "Text3": self.xl_reset_Text3[loop],
                "Text4": self.xl_reset_Text4[loop],
                "Text5": self.xl_reset_Text5[loop],
                "Text6": self.xl_reset_Text6[loop],
                "Text7": self.xl_reset_Text7[loop],
                "Text8": self.xl_reset_Text8[loop],
                "Text9": self.xl_reset_Text9[loop],
                "Text10": self.xl_reset_Text10[loop],
                "Text11": self.xl_reset_Text11[loop],
                "Text12": self.xl_reset_Text12[loop],
                "Text13": self.xl_reset_Text13[loop],
                "Text14": self.xl_reset_Text14[loop],
                "Text15": self.xl_reset_Text15[loop],
                "DateCustomField1": self.xl_reset_DateCustomField1[loop],
                "DateCustomField2": self.xl_reset_DateCustomField2[loop],
                "DateCustomField3": self.xl_reset_DateCustomField3[loop],
                "DateCustomField4": self.xl_reset_DateCustomField4[loop],
                "DateCustomField5": self.xl_reset_DateCustomField5[loop],
                "TextArea1": self.xl_reset_TextArea1[loop],
                "TextArea2": self.xl_reset_TextArea2[loop],
                "TextArea3": self.xl_reset_TextArea3[loop],
                "TextArea4": self.xl_reset_TextArea4[loop],
                "TrueFalse1": truefalse1,
                "TrueFalse2": truefalse2,
                "TrueFalse3": truefalse3,
                "TrueFalse4": truefalse4,
                "TrueFalse5": truefalse5
            },
            "CandidateId": self.xl_reset_candidate_id[loop]
        }
        reset_api = requests.post(api.web_api['update_candidate_details'],
                                  headers=self.headers, data=json.dumps(request, default=str), verify=False)
        print(reset_api.headers)
        reset_api_response = json.loads(reset_api.content)
        print(reset_api_response)


Object = ResetCandidateDetails()
Object.read_excel()
Total_count = len(Object.xl_reset_candidate_id)
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print("Iteration Count is ::", looping)
        Object.reset_candidate_details(looping)

        Object.headers = {}
