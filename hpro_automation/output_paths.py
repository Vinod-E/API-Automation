import os
# import datetime

# now = datetime.datetime.now()
# current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")

path = os.getenv("HOME")
generic_output_path = "%s/hirepro_automation/API-Automation/Output Data/" % path

# -------------------------
# Creation of Output sheets
# -------------------------
outputpaths = {

    # ------------------
    # CRPO Output sheets
    # ------------------
    'Password_policy': generic_output_path + "Crpo/PasswordPolicy/API_PasswordPolicy.xls",

    'CreateUser_Output_sheet': generic_output_path + "Crpo/CreateUser/API_CreateUser.xls",

    'UpdateUser_Output_sheet': generic_output_path + "Crpo/UpdateUser/API_UpdateUser.xls",

    'Score_Output_sheet': generic_output_path + "Crpo/UploadScore/API_Download_upload_Scores.xls",

    'Candidate_Output_sheet': generic_output_path + "Crpo/UploadCandidates/API_UploadCandidates.xls",

    'Update_Candidate_Output_sheet': generic_output_path + "Crpo/UpdateCandidates/API_Update_Candidates.xls",

    'SC_Output_sheet': generic_output_path + "Crpo/ShortlistingPanel/API_Shortlisting_Panel.xls",

    'EC_Output_sheet': generic_output_path + "Crpo/EC/API_DB_EC_Verification.xls",

    'Applicant_count_Output_sheet_1': generic_output_path + "Crpo/Event Applicant Search/"
                                                            "API_Event_Applicant_Search(Count).xls",

    'Applicant_count_Output_sheet_2': generic_output_path + "Crpo/Event Applicant Search/"
                                                            "API_Event_Applicant_BoundarySearch(Count).xls",

    'candidate_search_output_sheet_1': generic_output_path + "Crpo/Candidate Search/"
                                                             "API_Candidate_Search_With_Boundary.xls",

    'candidate_search_output_sheet_2': generic_output_path + "Crpo/Candidate Search/"
                                                             "API_Candidate_Search_Only_Count.xls",

    'Interview_flow_Output_sheet': generic_output_path + "Crpo/InterviewFlow/ProvideFeedback/API_ProvideFeedback.xls",

    'Reschdeule_Output_sheet': generic_output_path + "Crpo/InterviewFlow/Reschedule/API_Reschedule.xls",

    'Cancel_Interview_Output_sheet': generic_output_path + "Crpo/InterviewFlow/Cancel/API_Cancel_Interview.xls",

    'Activity_CallBack_Output_sheet': generic_output_path + "Crpo/Activity Call Back/API_Activity_CallBack.xls",

    'Duplication_rule_Output_sheet': generic_output_path + "Crpo/Duplication_rule/API_Duplication_rule.xls",

    'Login_check_Output_sheet': generic_output_path + "Crpo/Login/API_Login_check.xls",

    'Communication_output_sheet': generic_output_path + "Crpo/Communication_history/API_Communication_history.xls",

    'MI_output_sheet': generic_output_path + "Crpo/Manage Interviewers/API_Manage_Interviewers.xls",

    'Event_output_sheet': generic_output_path + "Crpo/Create_update_Event/API_Create_Update_Event.xls",

    # ------------------
    # Pofu output sheets
    # ------------------
    'p_candidates_output_sheet': generic_output_path + "Pofu/Upload_candidates/UploadCandidates.xls",


    # ------------------
    # Rpo output sheets
    # -----------------
    'r_Job_search_output_sheet': generic_output_path + "Rpo/Search/Combined_Job_Search.xls",

              }


# -------------------------
# DML scripts Output sheet
# -------------------------
DMLOutput = {'candidate_output_sheet': generic_output_path + 'Crpo/UploadCandidates/API_UploadCandidates.xls',
             'User_output_sheet': generic_output_path + "Crpo/CreateUser/API_CreateUser.xls",
             'Uploadscore_output_sheet': generic_output_path + "Crpo/UploadScore/API_Download_upload_Scores.xls"
             }
