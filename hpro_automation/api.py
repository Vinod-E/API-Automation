from hpro_automation.identity import credentials

generic_domain = credentials.generic_domain['domain']

# ---------------------------------------------Lambda APIS -------------------------------------------------------------
lambda_apis = {

    "bulkCreateTagCandidates": generic_domain + "crpo/candidate/api/v1/bulkCreateTagCandidates/",

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

    # ---------------------------------- communication history/status --------------------------------------------------
    "Create_Attachment": generic_domain + "common/attachment/api/v1/createAttachment/",

    "delete_Attachment": generic_domain + "common/attachment/api/v1/deleteAttachmentsForIds/",

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "gettaskbycandidate": generic_domain + "pofu/api/v1/get-task-by-candidate/",

}

# -------------------------------------------- Non Lambda APIS ---------------------------------------------------------
non_lambda_apis = {

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

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "submitform": generic_domain + "pofu/api/v1/submit-form/",

    "Approve_task": generic_domain + "pofu/api/v1/update-candidate-task-status/",

    "bulkimport": generic_domain + "pofu/api/v1/bulkimport",

}
