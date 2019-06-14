
generic_domain = "https://amsin.hirepro.in/"

web_api = {

    # ----------
    # CRPO APIS
    # ----------

    "create_update_pwd_policy": generic_domain + "py/common/user/create_update_pwd_policy/",

    "remove_pwd_policy": generic_domain + "py/common/user/remove_pwd_policy/",

    "change_password": generic_domain + "py/common/user/change_password/",

    "Loginto_CRPO": generic_domain + "py/common/user/login_user/",

    "bulkCreateTagCandidates": generic_domain + "py/crpo/candidate/api/v1/bulkCreateTagCandidates/",

    "CandidateGetbyId": generic_domain + "py/rpo/get_candidate_by_id/{}/",

    "Candidate_Educationaldetails": generic_domain + "py/rpo/get_candidate_education_details/{}/",

    "Candidate_ExperienceDetails": generic_domain + "py/rpo/get_candidate_experience_details/{}/",

    "update_candidate_details": 'https://amsin.hirepro.in/py/rpo/update_candidate_details/',

    "getAllApplicants": generic_domain + "py/crpo/applicant/api/v1/getAllApplicants/",

    "getAllEventApplicant": generic_domain + "py/crpo/applicant/api/v1/getAllEventApplicant/",

    "getApplicantsInfo": generic_domain + "py/crpo/applicant/api/v1/getApplicantsInfo/",

    "ChangeApplicant_Status": generic_domain + "py/crpo/applicant/api/v1/applicantStatusChange/",

    "createOrUpdateEcConfig": generic_domain + "py/crpo/dynamicec/api/v1/createOrUpdateEcConfig/",

    "Create_user": generic_domain + "py/common/user/create_user/",

    "Update_user": 'https://amsin.hirepro.in/py/common/user/update_user/',

    "UserGetByid": generic_domain + "py/common/user/get_user_by_id/{}/",

    "uploadCandidatesScore": generic_domain + "py/crpo/assessment/api/v1/uploadCandidatesScore/",

    "oneClickShortlist": generic_domain + "py/crpo/shortlistingcriteria/api/v1/oneClickShortlist",

    # ---------------------------------- communication history/status --------------------------------------------------
    "sendAdmitCardsToApplicants":
        generic_domain + "py/crpo/candidatecommunication/api/v1/sendAdmitCardsToApplicants",

    "sendRegistrationLinkToApplicants":
        generic_domain + "py/crpo/candidatecommunication/api/v1/sendRegistrationLinkToApplicants/",

    "Create_Attachment": generic_domain + "py/common/attachment/api/v1/createAttachment/",

    "delete_Attachment": generic_domain + "py/common/attachment/api/v1/deleteAttachmentsForIds/",

    "setApplicantCommunicationStatus":
        generic_domain + "py/crpo/candidatecommunication/api/v1/setApplicantCommunicationStatus",

    "sendVerificationNotification": generic_domain + "py/crpo/candidate/api/v1/sendVerificationNotification/",

    "getRegistrationLinkForApplicants":
        generic_domain + "py/crpo/candidatecommunication/api/v1/getRegistrationLinkForApplicants/",

    "applicantRe-Registration":
        generic_domain + "py/crpo/candidatecommunication/api/v1/applicantRe-Registration/",
    # ------------------------------------------------------------------------------------------------------------------

    "Schedule": generic_domain + "py/crpo/api/v1/interview/schedule/",

    "givefeedback": generic_domain + "py/crpo/api/v1/interview/givefeedback/",

    "Interview_details": generic_domain + "py/crpo/api/v1/interview/get/{}",

    "updateinterviewerdecision": generic_domain + "py/crpo/api/v1/interview/updateinterviewerdecision",

    "updateinterviewerfeedback": generic_domain + "py/crpo/api/v1/interview/updateinterviewerfeedback",

    "Reschedule": generic_domain + "py/crpo/api/v1/interview/reschedule/",

    "InterviewRequest_details": generic_domain + "py/crpo/api/v1/view/interviews",

    "cancel": generic_domain + "py/crpo/api/v1/interview/cancel/",

    "get_all_candidates": generic_domain + "py/rpo/get_all_candidates/",

    "gettaskbycandidate": generic_domain + "py/pofu/api/v1/get-task-by-candidate/",

    "submitform": generic_domain + "py/pofu/api/v1/submit-form/",

    "Approve_task": generic_domain + "py/pofu/api/v1/update-candidate-task-status/",

    # ----------
    # POFU APIS
    # ----------
    "bulkimport": generic_domain + "py/pofu/api/v1/bulkimport",

    # ----------
    # Rpo APIS
    # ----------
    "get_all_jobs": generic_domain + "py/rpo/get_all_jobs/"

           }

# ---------------------------------------------Lambda APIS -------------------------------------------------------------
lambda_apis = {

    "create_update_pwd_policy": "https://amsin.hirepro.in/py/common/user/create_update_pwd_policy/",

    "remove_pwd_policy": "https://amsin.hirepro.in/py/common/user/remove_pwd_policy/",

    "Loginto_CRPO": "https://amsin.hirepro.in/py/common/user/login_user/",

    "CandidateGetbyId": "https://amsin.hirepro.in/py/rpo/get_candidate_by_id/{}/",

    "get_all_jobs": "https://amsin.hirepro.in/py/rpo/get_all_jobs/",

    "Candidate_Educationaldetails": "https://amsin.hirepro.in/py/rpo/get_candidate_education_details/{}/",

    "Candidate_ExperienceDetails": "https://amsin.hirepro.in/py/rpo/get_candidate_experience_details/{}/",

    "Create_user": "https://amsin.hirepro.in/py/common/user/create_user/",

    "Update_user": 'https://amsin.hirepro.in/py/common/user/update_user/',

    "UserGetByid": "https://amsin.hirepro.in/py/common/user/get_user_by_id/{}/",

    # ---------------------------------- communication history/status --------------------------------------------------
    "Create_Attachment": 'https://amsin.hirepro.in/py/common/attachment/api/v1/createAttachment/',

    "delete_Attachment": 'https://amsin.hirepro.in/py/common/attachment/api/v1/deleteAttachmentsForIds/',

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "gettaskbycandidate": "https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",

}

# -------------------------------------------- Non Lambda APIS ---------------------------------------------------------
non_lambda_apis = {

    "change_password": "https://amsin.hirepro.in/py/common/user/change_password/",

    "bulkCreateTagCandidates": "https://amsin.hirepro.in/py/crpo/candidate/api/v1/bulkCreateTagCandidates/",

    "update_candidate_details": 'https://amsin.hirepro.in/py/rpo/update_candidate_details/',

    "getAllApplicants": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllApplicants/",

    "getAllEventApplicant": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllEventApplicant/",

    "getApplicantsInfo": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getApplicantsInfo/",

    "ChangeApplicant_Status": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/applicantStatusChange/",

    "createOrUpdateEcConfig": "https://amsin.hirepro.in/py/crpo/dynamicec/api/v1/createOrUpdateEcConfig/",

    "uploadCandidatesScore": "https://amsin.hirepro.in/py/crpo/assessment/api/v1/uploadCandidatesScore/",

    "oneClickShortlist": "https://amsin.hirepro.in/py/crpo/shortlistingcriteria/api/v1/oneClickShortlist",

    # ---------------------------------- communication history/status --------------------------------------------------
    "sendAdmitCardsToApplicants":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/sendAdmitCardsToApplicants",

    "sendRegistrationLinkToApplicants":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/sendRegistrationLinkToApplicants/",

    "setApplicantCommunicationStatus":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/setApplicantCommunicationStatus",

    "sendVerificationNotification": "https://amsin.hirepro.in/py/crpo/candidate/api/v1/sendVerificationNotification/",

    "getRegistrationLinkForApplicants":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/getRegistrationLinkForApplicants/",

    "applicantRe-Registration":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/applicantRe-Registration/",

    # ------------------------------------------- Interview ------------------------------------------------------------
    "Schedule": "https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",

    "givefeedback": "https://amsin.hirepro.in/py/crpo/api/v1/interview/givefeedback/",

    "Interview_details": "https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}",

    "updateinterviewerdecision": "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerdecision",

    "updateinterviewerfeedback": "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerfeedback",

    "Reschedule": "https://amsin.hirepro.in/py/crpo/api/v1/interview/reschedule/",

    "InterviewRequest_details": "https://amsin.hirepro.in/py/crpo/api/v1/view/interviews",

    "cancel": "https://amsin.hirepro.in/py/crpo/api/v1/interview/cancel/",

    # --------------------------------------------- Staffing -----------------------------------------------------------
    "submitform": "https://amsin.hirepro.in/py/pofu/api/v1/submit-form/",

    "Approve_task": "https://amsin.hirepro.in/py/pofu/api/v1/update-candidate-task-status/",

    "bulkimport": "https://amsin.hirepro.in/py/pofu/api/v1/bulkimport",

}
