import os
path = os.getenv("HOME")

generic_input_path = "%s/hirepro_automation/API-Automation/Input Data/" % path

Uploadscore_Input_sheet = generic_input_path + "CreateUser/CreateUser.xls"

inputpaths = {

    # ------------
    # CRPO Inputs
    # ------------
    'PasswordPolicy_Input_sheet': generic_input_path + "Crpo/PasswordPolicy/Password Policy.xls",

    'CreateUser_Input_sheet': generic_input_path + "Crpo/CreateUser/CreateUser.xls",

    'updateUser_Input_sheet': generic_input_path + "Crpo/Update_user/update_user.xls",

    'resetUser_Input_sheet': generic_input_path + "Crpo/Update_user/reset_user.xls",

    'Uploadscore_Input_sheet': generic_input_path + 'Crpo/ScoreSheet/UploadScores.xls',
    'Group_Input_sheet': generic_input_path + "Crpo/ScoreSheet/Group.xlsx",
    'Group_Section_Input_sheet': generic_input_path + "Crpo/ScoreSheet/Group_Section.xlsx",
    'Updated_group_Input_sheet': generic_input_path + "Crpo/ScoreSheet/Updated_Group.xlsx",
    'Update_group_section_input_sheet': generic_input_path + "Crpo/ScoreSheet/Updated_Group_Section.xlsx",

    'UploadCandi_Input_sheet': generic_input_path + "Crpo/UploadCandidates/Candidate_Upload.xls",

    'Update_candi_Input_sheet': generic_input_path + 'Crpo/Update_candidate/Update_candidates.xls',

    'reset_candi_Input_sheet': generic_input_path + 'Crpo/Update_candidate/reset_candidates.xls',

    'EC_Input_sheet': generic_input_path + "Crpo/EC/Final_EC_Configuration.xls",

    'Shortlisting_Input_sheet': generic_input_path + "Crpo/ShortlistingPanel/Shortlisting_Panel.xls",

    'Communication_History_Input_sheet': generic_input_path + "Crpo/Communication/Communication.xls",

    'Provide_feedback_Input_sheet': generic_input_path + "Crpo/InterviewFlow/ProvideFeedback/GiveFeedback.xls",

    'Reschedule_input_sheet': generic_input_path + "Crpo/InterviewFlow/Reschedule_Cancel/Reschedule.xls",

    'Cancel_interview_input_sheet': generic_input_path + "Crpo/InterviewFlow/Cancel/CancelInterview.xls",

    'Applicant_count_Input_sheet': generic_input_path + "Crpo/Event Applicant Search/"
                                                        "Event_applicant_Search_Boundary_Condition.xls",
    'candidate_search_Input_sheet': generic_input_path + "Crpo/Candidate search/"
                                                         "Candidate_Combined_Search_Boundary_Condition.xls",
    'Activity_C_back_Input_sheet': generic_input_path + "Crpo/Activity Call Back/Activitycallback.xls",

    'Duplication_rule_Input_sheet': generic_input_path + "Crpo/Duplication_rule/Duplication_Rule.xls",

    'Login_check_Input_sheet': generic_input_path + "Crpo/Login/Login.xls",

    'Manage_Int_Input_sheet': generic_input_path + "Crpo/Manage Interviewers/Manage_Interviewers.xls",

    'Event_Input_sheet': generic_input_path + "Crpo/Create_update_Event/Create_Update_Event.xls",

    'CloneEvent_Input_sheet': generic_input_path + "Crpo/Clone Event/Clone Event.xls",

    # ------------
    # Pofu Inputs
    # ------------
    'p_upload_candidates_Input_sheet': generic_input_path + "Pofu/Upload_candidates/UploadExcel.xls",


    # -----------
    # Rpo Inputs
    # -----------
    'r_job_search_Input_sheet': generic_input_path + "Rpo/Search/JobSearchWithIds.xls"

}

driver = {
    'chrome': '%s/hirepro_automation/API-Automation/Utilities/chromedriver' % path
}
