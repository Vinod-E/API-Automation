import requests
import json


class AllConfigurations:

    def __init__(self):
        requests.packages.urllib3.disable_warnings()
        self.stack_ranking_report_payload = {}

    def apiLists(self):
        self.login = "https://amsin.hirepro.in/py/common/user/login_user/"
        self.getall_applicant_api = "https://amsin.hirepro.in/py/common/xl_creator/api/v1/generate_applicant_report/"

    def filePath(self, date):
        self.expected_excel_sheet_path = '/home/vinod/hirepro_automation/API-Automation/Input Data/' \
                                         'Crpo/stackranking/stack_ranking_FourEvents_report.xlsx'

        self.download_path = '/home/vinod/hirepro_automation/API-Automation/Output Data/Crpo/stackranking/Downloads/' \
                             'stackranking_%s.xlsx' % date

        self.save_path = '/home/vinod/hirepro_automation/API-Automation/Output Data/Crpo/stackranking/' \
                         'stackranking_op_file%s.xlsx' % date


    def loginToCRPO(self):
        header = {"content-type": "application/json"}
        data = {"LoginName": "admin", "Password": "4LWS-067", "TenantAlias": "automation", "UserName": "admin"}
        response = requests.post(self.login, headers=header,
                                 data=json.dumps(data), verify=False)
        login_response = response.json()
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": login_response.get("Token")}

    # ------------------------------------------------------------------------------------------------------------------#
    # 1. This method is having all api requests
    # 2. Here, all requests are used for download Stack Ranking report API
    # ------------------------------------------------------------------------------------------------------------------#
    def apiRequests(self):
        self.stack_ranking_report_payload = {"applicantIds": [896910, 896909, 896911, 896908],
                                             "templateJSON": {"format_rules": {
                                                 "header_format": {"bold": 1, "border": 1, "border_color": "#000000",
                                                                   "align": "center", "valign": "vcenter",
                                                                   "color": "#ffffff", "fg_color": "#006fa6"},
                                                 "align_center": {"align": "center", "valign": "vcenter"},
                                                 "date": {"align": "center", "valign": "vcenter",
                                                          "num_format": "DD MMMM YYYY"},
                                                 "dateTime": {"align": "center", "valign": "vcenter",
                                                              "num_format": "HH:mm:ss DD MMMM YYYY"},
                                                 "wrap_text": {"align": "center", "valign": "vcenter",
                                                               "text_wrap": True}},
                                                 "transform_rules": ["SPLIT_ED_PROFILE"],
                                                 "file_meta_rules": {
                                                     "file_name": "stack_ranking_FourEvents_report"},
                                                 "content_rules": [{"columnName": "Applicant Id",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "applicantDetail",
                                                                        "applicantId"],
                                                                    "default": " ",
                                                                    "id": "ApplicantId"},
                                                                   {"columnName": "Candidate Id",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "candidateId"],
                                                                    "default": " ",
                                                                    "id": "CandidateId"}, {
                                                                       "columnName": "Candidate Name",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "fullName"],
                                                                       "default": " ",
                                                                       "id": "Name"}, {
                                                                       "columnName": "Primary Email",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "primaryEmail"],
                                                                       "default": " ",
                                                                       "id": "Email1"},
                                                                   {"columnName": "Mobile",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "mobile"],
                                                                    "default": " ",
                                                                    "id": "Mobile1"},
                                                                   {"columnName": "Gender",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "gender"],
                                                                    "default": " ",
                                                                    "id": "Gender"}, {
                                                                       "columnName": "Date Of Birth",
                                                                       "format": "date",
                                                                       "is_date": True,
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "dateOfBirth"],
                                                                       "default": " ",
                                                                       "id": "DateOfBirth"}, {
                                                                       "columnName": "Current Stage",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "applicantDetail",
                                                                           "currentStageLabel"],
                                                                       "default": " ",
                                                                       "id": "CurrentStage"}, {
                                                                       "columnName": "Current Status",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "applicantDetail",
                                                                           "currentStatusLabel"],
                                                                       "default": " ",
                                                                       "id": "CurrentStatus"},
                                                                   {"columnName": "Degree",
                                                                    "format": "align_center",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "SPLIT_ED_PROFILE", "PG",
                                                                        "degree"],
                                                                    "default": " ",
                                                                    "id": "EducationProfiles"},
                                                                   {"columnName": "Percentage",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "SPLIT_ED_PROFILE",
                                                                        "Final",
                                                                        "percentageOrCgp"],
                                                                    "default": " ",
                                                                    "id": "EducationProfiles"}, {
                                                                       "columnName": "Percentage Out Of",
                                                                       "format": "align_center",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "SPLIT_ED_PROFILE",
                                                                           "Final",
                                                                           "percentageOutOf"],
                                                                       "default": " ",
                                                                       "id": "EducationProfiles"},
                                                                   {
                                                                       "columnName": "Year of Passing",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "SPLIT_ED_PROFILE",
                                                                           "Final",
                                                                           "yearOfPassing"],
                                                                       "default": " ",
                                                                       "id": "EducationProfiles"},
                                                                   {"columnName": "Branch",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "SPLIT_ED_PROFILE",
                                                                        "Final", "branch"],
                                                                    "default": " ",
                                                                    "id": "EducationProfiles"},
                                                                   {"columnName": "College",
                                                                    "format": "align_center",
                                                                    "force_inclusion": "True",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "SPLIT_ED_PROFILE",
                                                                        "Final", "college"],
                                                                    "default": " ",
                                                                    "id": "EducationProfiles"},
                                                                   {"columnName": "University",
                                                                    "format": "align_center",
                                                                    "accessor": [
                                                                        "candidateDetails",
                                                                        "SPLIT_ED_PROFILE",
                                                                        "Final", "university"],
                                                                    "default": " ",
                                                                    "id": "EducationProfiles"}, {
                                                                       "columnName": "Education City",
                                                                       "format": "align_center",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "SPLIT_ED_PROFILE",
                                                                           "Final", "city"],
                                                                       "default": " ",
                                                                       "id": "EducationProfiles"},
                                                                   {
                                                                       "columnName": "Education State",
                                                                       "format": "align_center",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "SPLIT_ED_PROFILE",
                                                                           "Final", "state"],
                                                                       "default": " ",
                                                                       "id": "EducationProfiles"},
                                                                   {
                                                                       "columnName": "Gap in Education",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "customDetails",
                                                                           "integer1",
                                                                           "propertyValue"],
                                                                       "default": " ",
                                                                       "override_column_length": 18,
                                                                       "id": "Integer1"}, {
                                                                       "columnName": "Standing Backlog",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "customDetails",
                                                                           "integer3",
                                                                           "propertyValue"],
                                                                       "default": " ",
                                                                       "override_column_length": 18,
                                                                       "id": "Integer3"}, {
                                                                       "columnName": "Have you applied for Accenture in last 6 months (excluding current event)",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "customDetails",
                                                                           "integer4",
                                                                           "propertyValue"],
                                                                       "default": " ",
                                                                       "override_column_length": 75,
                                                                       "id": "Integer4"}, {
                                                                       "columnName": "Are you an Indian citizen?",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "customDetails",
                                                                           "integer7",
                                                                           "propertyValue"],
                                                                       "default": " ",
                                                                       "override_column_length": 28,
                                                                       "id": "Integer7"}, {
                                                                       "columnName": "Have you worked in Accenture before",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": [
                                                                           "candidateDetails",
                                                                           "customDetails",
                                                                           "integer2",
                                                                           "propertyValue"],
                                                                       "default": " ",
                                                                       "override_column_length": 37,
                                                                       "id": "Integer2"}, {
                                                                       "columnName": "Stack Percentage",
                                                                       "format": "align_center",
                                                                       "force_inclusion": "True",
                                                                       "accessor": "stackPercentage",
                                                                       "default": " ",
                                                                       "id": "stackPercentage"},
                                                                   {
                                                                       "columnName": "Stack Ranking",
                                                                       "accessor": "scoreDetails",
                                                                       "default": " ",
                                                                       "include_child_enveloping_header": "True",
                                                                       "include_child_indexed_header": "True",
                                                                       "collate_by_key": "groups",
                                                                       "next": [{
                                                                           "columnName": "weightage",
                                                                           "format": "align_center",
                                                                           "accessor": "weightage",
                                                                           "default": " "},
                                                                           {
                                                                               "columnName": "percentage",
                                                                               "format": "align_center",
                                                                               "accessor": "percentage",
                                                                               "default": " "},
                                                                           {
                                                                               "columnName": "score",
                                                                               "format": "align_center",
                                                                               "accessor": "score",
                                                                               "default": " "},
                                                                           {
                                                                               "columnName": "maxScore",
                                                                               "format": "align_center",
                                                                               "accessor": "maxScore",
                                                                               "default": " "}]}]},
                                             "invokeSync": True, "jsonToMerge": {"key_name": "applicantId", "data": [
                {"stackRanking": 1, "scoreDetails": [
                    {"weightage": 30, "percentage": 21, "score": 7, "maxScore": 10,
                     "groups": "Aptitude Technical Automobile"},
                    {"weightage": 25, "percentage": 23.75, "score": 19, "maxScore": 20,
                     "groups": "HRM General Knowledge"}],
                 "applicantId": 896910, "groups": {
                    "HRM": {"totalMarks": 10, "obtainMarks": 10, "testId": 8999, "attendedOn": "2020-01-08 14:32:42",
                            "partnerIntegrationId": 21},
                    "Automobile": {"totalMarks": 10, "obtainMarks": 7, "testId": 8999,
                                   "attendedOn": "2020-01-08 14:32:42",
                                   "partnerIntegrationId": 21},
                    "General Knowledge": {"totalMarks": 10, "obtainMarks": 9, "testId": 8999,
                                          "attendedOn": "2020-01-08 14:32:42", "partnerIntegrationId": 21}},
                 "stackPercentage": 44.75},
                {"stackRanking": 2, "scoreDetails": [
                    {"weightage": 30, "percentage": 22.5,
                     "score": 15, "maxScore": 20,
                     "groups": "Aptitude Technical Automobile"},
                    {"weightage": 15, "percentage": 13.5,
                     "score": 9, "maxScore": 10,
                     "groups": "Coding Assessment"}],
                 "applicantId": 896909, "groups": {
                    "Technical": {"totalMarks": 10,
                                  "obtainMarks": 8,
                                  "testId": 8998,
                                  "attendedOn": "2020-01-08 14:31:53",
                                  "partnerIntegrationId": 22},
                    "Coding Assessment": {"totalMarks": 10,
                                          "obtainMarks": 9,
                                          "testId": 8998,
                                          "attendedOn": "2020-01-08 14:31:53",
                                          "partnerIntegrationId": 22},
                    "Aptitude": {"totalMarks": 10,
                                 "obtainMarks": 7,
                                 "testId": 8998,
                                 "attendedOn": "2020-01-08 14:31:53",
                                 "partnerIntegrationId": 22}},
                 "stackPercentage": 36},
                {"stackRanking": 3, "scoreDetails": [
                    {"weightage": 30, "percentage": 30,
                     "score": 30, "maxScore": 30,
                     "groups": "HRMM GK Automobile Engineering"}],
                 "applicantId": 896911, "groups": {
                    "HRMM": {"totalMarks": 10,
                             "obtainMarks": 10,
                             "testId": 9000,
                             "attendedOn": "2020-01-08 14:34:17",
                             "partnerIntegrationId": 23},
                    "Automobile Engineering": {
                        "totalMarks": 10,
                        "obtainMarks": 10, "testId": 9000,
                        "attendedOn": "2020-01-08 14:34:17",
                        "partnerIntegrationId": 23},
                    "GK": {"totalMarks": 10,
                           "obtainMarks": 10,
                           "testId": 9000,
                           "attendedOn": "2020-01-08 14:34:17",
                           "partnerIntegrationId": 23}},
                 "stackPercentage": 30},
                {"stackRanking": 4, "scoreDetails": [
                    {"weightage": 30, "percentage": 18,
                     "score": 18, "maxScore": 30,
                     "groups": "Aptitude Technical Automobile"}],
                 "applicantId": 896908, "groups": {
                    "Technical": {"totalMarks": 10,
                                  "obtainMarks": 6,
                                  "testId": 8997,
                                  "attendedOn": "2020-01-08 14:30:49",
                                  "partnerIntegrationId": 20},
                    "Automobile": {"totalMarks": 10,
                                   "obtainMarks": 7,
                                   "testId": 8997,
                                   "attendedOn": "2020-01-08 14:30:49",
                                   "partnerIntegrationId": 20},
                    "Aptitude": {"totalMarks": 10,
                                 "obtainMarks": 5,
                                 "testId": 8997,
                                 "attendedOn": "2020-01-08 14:30:49",
                                 "partnerIntegrationId": 20}},
                 "stackPercentage": 18}]}}


config_obj = AllConfigurations()
config_obj.apiLists()
config_obj.loginToCRPO()
config_obj.apiRequests()
