from hpro_automation.identity import credentials

login_server = str(input('Server_name :: '))
generic_domain = credentials.server_api.format(login_server)
eu_ams_domain = credentials.ams_eu_server_api
eu_amsin_domain = credentials.amsin_eu_server_api
# print('That you are running in this server :: ', generic_domain)
# ---------------------------------------------Lambda APIS -------------------------------------------------------------
lambda_apis = {

    "getAllEvent": generic_domain + 'crpo/event/api/v1/getAllEvent/',

    "getAllApplicants": generic_domain + "crpo/applicant/api/v1/getAllApplicants/",

    "create_update_pwd_policy": generic_domain + "common/user/create_update_pwd_policy/",

    "remove_pwd_policy": generic_domain + "common/user/remove_pwd_policy/",

    "Loginto_CRPO": generic_domain + "common/user/login_user/",

    "CandidateGetbyId": generic_domain + "rpo/get_candidate_by_id/{}/",

    "get_all_jobs": generic_domain + "rpo/get_all_jobs/",

    "Candidate_Educationaldetails": generic_domain + "rpo/get_candidate_education_details/{}/",

    "Candidate_ExperienceDetails": generic_domain + "rpo/get_candidate_experience_details/{}/",

    "Create_user": generic_domain + "common/user/create_user/",

    "Update_user": generic_domain + "common/user/update_user/",

    "UserGetByid": generic_domain + "common/user/get_user_by_id/{}/",

    "save_app_preferences": generic_domain + "common/common_app_utils/save_app_preferences/",

    "getAllAppPreference": generic_domain + 'common/common_app_utils/api/v1/getAllAppPreference/',

    "getPartialGetEventForId": generic_domain + "crpo/event/api/v1/getPartialGetEventForId/",

    "getEcConfigs": generic_domain + "crpo/dynamicec/api/v1/getEcConfigs/",

    "verfiyhash": generic_domain + "crpo/assessment/slotmgmt/candidate/api/v1/verifyHash/",

    "generate_applicant_report": generic_domain + "common/xl_creator/api/v1/generate_applicant_report/",

    "assessment_slot_select": generic_domain + "crpo/assessment/slotmgmt/candidate/api/v1/selectSlot/",

    "assessment_slot_update": generic_domain + "crpo/assessment/slotmgmt/candidate/api/v1/changeSlot/",

    "applicant_screen_data": generic_domain + "crpo/candidate/api/v1/getScreenData/",

    # ---------------------------------- communication history/status --------------------------------------------------
    "Create_Attachment": generic_domain + "common/attachment/api/v1/createAttachment/",

    "delete_Attachment": generic_domain + "common/attachment/api/v1/deleteAttachmentsForIds/",

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "gettaskbycandidate": generic_domain + "pofu/api/v1/get-task-by-candidate/",

    # ------------------------------------------- Performance APIS -----------------------------------------------------
    "get_tenant_details": generic_domain + "common/get_tenant_details/",
    "get_all_entity_properties": generic_domain + "rpo/get_all_entity_properties/",
    "group_by_catalog_masters": generic_domain + "common/catalogs/api/v1/group-by-catalog-masters/",
    "get_all_candidates": generic_domain + "rpo/get_all_candidates/",
    "getTestUsersForTest": generic_domain + "assessment/testuser/api/v1/getTestUsersForTest/",
    "interviews": generic_domain + "crpo/api/v1/view/interviews",
    "interview_new": generic_domain + "crpo/api/v1/view/interviewsNew",

    "eu_amsin_login": eu_amsin_domain + "common/user/login_user/",
    "amsin_eu_getAllAppPreference": eu_amsin_domain + 'common/common_app_utils/api/v1/getAllAppPreference/',
    "amsin_eu_get_tenant_details": eu_amsin_domain + "common/get_tenant_details/",
    "amsin_eu_get_all_entity_properties": eu_amsin_domain + "rpo/get_all_entity_properties/",
    "amsin_eu_group_by_catalog_masters": eu_amsin_domain + "common/catalogs/api/v1/group-by-catalog-masters/",
    "amsin_eu_get_all_candidates": eu_amsin_domain + "rpo/get_all_candidates/",
    "amsin_eu_getTestUsersForTest": eu_amsin_domain + "assessment/testuser/api/v1/getTestUsersForTest/",
    "amsin_eu_interviews": eu_amsin_domain + "crpo/api/v1/view/interviews",
    "amsin_eu_interview_new": eu_amsin_domain + "crpo/api/v1/view/interviewsNew",

    "eu_ams_login": eu_ams_domain + "common/user/login_user/",
    "ams_eu_getAllAppPreference": eu_ams_domain + 'common/common_app_utils/api/v1/getAllAppPreference/',
    "ams_eu_get_tenant_details": eu_ams_domain + "common/get_tenant_details/",
    "ams_eu_get_all_entity_properties": eu_ams_domain + "rpo/get_all_entity_properties/",
    "ams_eu_group_by_catalog_masters": eu_ams_domain + "common/catalogs/api/v1/group-by-catalog-masters/",
    "ams_eu_get_all_candidates": eu_ams_domain + "rpo/get_all_candidates/",
    "ams_eu_getTestUsersForTest": eu_ams_domain + "assessment/testuser/api/v1/getTestUsersForTest/",
    "ams_eu_interviews": eu_ams_domain + "crpo/api/v1/view/interviews",
    "ams_eu_interview_new": eu_ams_domain + "crpo/api/v1/view/interviewsNew",

}

# -------------------------------------------- Non Lambda APIS ---------------------------------------------------------
non_lambda_apis = {

    "bulkCreateTagCandidates": generic_domain + "crpo/candidate/api/v1/bulkCreateTagCandidates/",

    "createEvent": generic_domain + "crpo/event/api/v1/createEvent/",

    "cloneEvent": generic_domain + "crpo/event/api/v1/cloneEvent/",

    "updateEvent": generic_domain + "crpo/event/api/v1/updateEvent/",

    "change_password": generic_domain + "common/user/change_password/",

    "update_candidate_details": generic_domain + "rpo/update_candidate_details/",

    "getAllEventApplicant": generic_domain + "crpo/applicant/api/v1/getAllEventApplicant/",

    "getApplicantsInfo": generic_domain + "crpo/applicant/api/v1/getApplicantsInfo/",

    "ChangeApplicant_Status": generic_domain + "crpo/applicant/api/v1/applicantStatusChange/",

    "createOrUpdateEcConfig": generic_domain + "crpo/dynamicec/api/v1/createOrUpdateEcConfig/",

    "uploadCandidatesScore": generic_domain + "crpo/assessment/api/v1/uploadCandidatesScore/",

    "oneClickShortlist": generic_domain + "crpo/shortlistingcriteria/api/v1/oneClickShortlist",

    "candidate_duplicate_check": generic_domain + "rpo/candidate_duplicate_check/",

    "get_all_candidates": generic_domain + "rpo/get_all_candidates/",

    "set_interviewer_nomination": generic_domain + "crpo/interviewer_nomination/api/v1/"
                                                   "set_interviewer_nomination_details_for_event/",

    "send_nomination_mails_to_selected_interviewers": generic_domain + "crpo/interviewer_nomination/api/v1/"
                                                                       "send_nomination_mails_to_"
                                                                       "selected_interviewers/",

    "update_interviewer_nomination_status": generic_domain + "crpo/interviewer_nomination/api/v1/"
                                                             "update_interviewer_nomination_status/",

    "get_all_invited_interviewers": generic_domain + "crpo/interviewer_nomination/api/v1/get_all_invited_interviewers/",

    "sync_interviewers": generic_domain + "crpo/interviewer_nomination/api/v1/sync_interviewers/",

    "get_event_wise_nominations_summary_count": generic_domain + "crpo/interviewer_nomination/api/v1/"
                                                                 "get_event_wise_nominations_summary_count/",

    "getAssessmentSummary": generic_domain + "crpo/assessment/api/v1/getAssessmentSummary/",

    "getEventRegistrationDates": generic_domain + "crpo/api/v1/getEventRegistrationDates",

    "authenticate": generic_domain + "crpo/lip/candidate/api/v1/authenticate/",

    "preference": generic_domain + "crpo/applicant/preference/api/v1/validate",

    "save_slot": generic_domain + "crpo/applicant/preference/api/v1/save",

    "interview_slot_select": generic_domain + 'crpo/lip/slot/api/v1/mark_applicants_slot/',

    # ---------------------------------- communication history/status --------------------------------------------------
    "sendAdmitCardsToApplicants":
        generic_domain + "crpo/candidatecommunication/api/v1/sendAdmitCardsToApplicants",

    "sendRegistrationLinkToApplicants":
        generic_domain + "crpo/candidatecommunication/api/v1/sendRegistrationLinkToApplicants/",

    "setApplicantCommunicationStatus":
        generic_domain + "crpo/candidatecommunication/api/v1/setApplicantCommunicationStatus",

    "sendVerificationNotification": generic_domain + "crpo/candidate/api/v1/sendVerificationNotification/",

    "getRegistrationLinkForApplicants":
        generic_domain + "crpo/candidatecommunication/api/v1/getRegistrationLinkForApplicants/",

    "applicantRe-Registration":
        generic_domain + "crpo/candidatecommunication/api/v1/applicantRe-Registration/",

    # ------------------------------------------- Interview ------------------------------------------------------------
    "Schedule": generic_domain + "crpo/api/v1/interview/schedule/",

    "givefeedback": generic_domain + "crpo/api/v1/interview/givefeedback/",

    "Interview_details": generic_domain + "crpo/api/v1/interview/get/{}",

    "updateinterviewerdecision": generic_domain + "crpo/api/v1/interview/updateinterviewerdecision",

    "updateinterviewerfeedback": generic_domain + "crpo/api/v1/interview/updateinterviewerfeedback",

    "Reschedule": generic_domain + "crpo/api/v1/interview/reschedule/",

    "InterviewRequest_details": generic_domain + "crpo/api/v1/view/interviews",

    "cancel": generic_domain + "crpo/api/v1/interview/cancel/",

    "interview_unassign_slot": generic_domain + "crpo/applicant/api/v1/unslotApplicants/",

    "assessment_unassign_slot": generic_domain + "crpo/assessment/slotmgmt/recruiter/api/v1/dissociateSlotApplicants/",

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "submitform": generic_domain + "pofu/api/v1/submit-form/",

    "Approve_task": generic_domain + "pofu/api/v1/update-candidate-task-status/",

    "bulkimport": generic_domain + "pofu/api/v1/bulkimport",

    "tenant_cache": generic_domain + "common/api/v1/ctic/",

    # ------------------------------------------- Performance APIS -----------------------------------------------------
    "api1": generic_domain + "common/get_tenant_details/",

    # ------------------------------------------- S3 APIS -----------------------------------------------------
    "s3": generic_domain + "common/filehandler/api/v2/upload/.jpeg,.jpg,.gif,.png,.pdf,.txt,.doc,.docx,.xls,.xlsx,.zip,"
                           ".rar,.7z,.msg,.html,.ogg,.mp4,.webm,/15000/",
}

slot_app = {
    # ------------------------------------------- Slot App ------------------------------------------------------------
    'access_token': generic_domain + "oauth2/{}/access_token/"
}
