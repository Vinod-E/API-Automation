import os
# import datetime

# now = datetime.datetime.now()
# current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")

path = os.getenv("HOME")
generic_output_path = "%s/hirepro_automation/API-Automation/Output Data/" % path
crpo_common_folder = "Crpo/Common_folder/"
crpo_DML_folder = "Crpo/DML_folder/"

# -------------------------
# Creation of Output sheets
# -------------------------
outputpaths = {

    # ------------------
    # CRPO Output sheets
    # ------------------
    'Password_policy': generic_output_path + crpo_common_folder + "API_PasswordPolicy.xls",

    'CreateUser_Output_sheet': generic_output_path + crpo_common_folder + "API_CreateUser.xls",

    'UpdateUser_Output_sheet': generic_output_path + crpo_common_folder + "API_UpdateUser.xls",

    'Score_Output_sheet': generic_output_path + crpo_common_folder + "API_Download_upload_Scores.xls",

    'Candidate_Output_sheet': generic_output_path + crpo_common_folder + "API_UploadCandidates.xls",

    'Update_Candidate_Output_sheet': generic_output_path + crpo_common_folder + "API_Update_Candidates.xls",

    'SC_Output_sheet': generic_output_path + crpo_common_folder + "API_Shortlisting_Panel.xls",

    'EC_Output_sheet': generic_output_path + crpo_common_folder + "API_DB_EC_Verification.xls",

    'Applicant_count_Output_sheet_1': generic_output_path + crpo_common_folder + "API_Event_Applicant_"
                                                                                 "Search(Count).xls",

    'Applicant_count_Output_sheet_2': generic_output_path + crpo_common_folder + "API_Event_Applicant_"
                                                                                 "BoundarySearch(Count).xls",

    'candidate_search_output_sheet_1': generic_output_path + crpo_common_folder + "API_Candidate_Search_With_"
                                                                                  "Boundary.xls",

    'candidate_search_output_sheet_2': generic_output_path + crpo_common_folder + "API_Candidate_Search_Only_Count.xls",

    'Interview_flow_Output_sheet': generic_output_path + crpo_common_folder + "API_ProvideFeedback.xls",

    'Reschdeule_Output_sheet': generic_output_path + crpo_common_folder + "API_Reschedule.xls",

    'Cancel_Interview_Output_sheet': generic_output_path + crpo_common_folder + "API_Cancel_Interview.xls",

    'Activity_CallBack_Output_sheet': generic_output_path + crpo_common_folder + "API_Activity_CallBack.xls",

    'Duplication_rule_Output_sheet': generic_output_path + crpo_common_folder + "API_Duplication_rule.xls",

    'Login_check_Output_sheet': generic_output_path + crpo_common_folder + "API_Login_check.xls",

    'Communication_output_sheet': generic_output_path + crpo_common_folder + "API_Communication_history.xls",

    'MI_output_sheet': generic_output_path + crpo_common_folder + "API_Manage_Interviewers.xls",

    'Event_output_sheet': generic_output_path + crpo_common_folder + "API_Create_Update_Event.xls",

    'Event_Clone_output_sheet': generic_output_path + crpo_common_folder + "API_Event_Clone.xls",

    'Stack_Ranking_download_sheet': generic_output_path + crpo_common_folder + "Downloads/stackranking_{}.xlsx",

    'Stack_Ranking_output_sheet': generic_output_path + crpo_common_folder + "stackranking_{}.xlsx",

    # ------------------
    # Pofu output sheets
    # ------------------
    'p_candidates_output_sheet': generic_output_path + "Pofu/Upload_candidates/UploadCandidates.xls",


    # ------------------
    # Rpo output sheets
    # -----------------
    'r_Job_search_output_sheet': generic_output_path + "Rpo/Search/Combined_Job_Search.xls",

    # ------ Performance -------
    'performance_testing': generic_output_path + crpo_common_folder + "Performance_testing.xlsx"

}
