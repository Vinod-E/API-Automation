web_api = {

    # ----------
    # CRPO APIS
    # ----------

    "create_update_pwd_policy": "https://amsin.hirepro.in/py/common/user/create_update_pwd_policy/",

    "remove_pwd_policy": "https://amsin.hirepro.in/py/common/user/remove_pwd_policy/",

    "change_password": "https://amsin.hirepro.in/py/common/user/change_password/",

    "Loginto_CRPO": "https://amsin.hirepro.in/py/common/user/login_user/",

    "bulkCreateTagCandidates": "https://amsin.hirepro.in/py/crpo/candidate/api/v1/bulkCreateTagCandidates/",

    "CandidateGetbyId": "https://amsin.hirepro.in/py/rpo/get_candidate_by_id/{}/",

    "Candidate_Educationaldetails": "https://amsin.hirepro.in/py/rpo/get_candidate_education_details/{}/",

    "Candidate_ExperienceDetails": "https://amsin.hirepro.in/py/rpo/get_candidate_experience_details/{}/",

    "update_candidate_details": 'https://amsin.hirepro.in/py/rpo/update_candidate_details/',

    "getAllApplicants": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllApplicants/",

    "getAllEventApplicant": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllEventApplicant/",

    "getApplicantsInfo": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/getApplicantsInfo/",

    "ChangeApplicant_Status": "https://amsin.hirepro.in/py/crpo/applicant/api/v1/applicantStatusChange/",

    "createOrUpdateEcConfig": "https://amsin.hirepro.in/py/crpo/dynamicec/api/v1/createOrUpdateEcConfig/",

    "Create_user": "https://amsin.hirepro.in/py/common/user/create_user/",

    "Update_user": 'https://amsin.hirepro.in/py/common/user/update_user/',

    "UserGetByid": "https://amsin.hirepro.in/py/common/user/get_user_by_id/{}/",

    "uploadCandidatesScore": "https://amsin.hirepro.in/py/crpo/assessment/api/v1/uploadCandidatesScore/",

    "oneClickShortlist": "https://amsin.hirepro.in/py/crpo/shortlistingcriteria/api/v1/oneClickShortlist",

    # ---------------------------------- communication history/status --------------------------------------------------
    "sendAdmitCardsToApplicants":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/sendAdmitCardsToApplicants",

    "sendRegistrationLinkToApplicants":
        "https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/sendRegistrationLinkToApplicants/",

    "Create_Attachment": 'https://amsin.hirepro.in/py/common/attachment/api/v1/createAttachment/',

    "delete_Attachment": 'https://amsin.hirepro.in/py/common/attachment/api/v1/deleteAttachmentsForIds/',

    "setApplicantCommunicationStatus": 'https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/s'
                                       'etApplicantCommunicationStatus',

    "sendVerificationNotification": 'https://amsin.hirepro.in/py/crpo/candidate/api/v1/sendVerificationNotification/',

    "getRegistrationLinkForApplicants": 'https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/'
                                        'getRegistrationLinkForApplicants/',

    "applicantRe-Registration": 'https://amsin.hirepro.in/py/crpo/candidatecommunication/api'
                                '/v1/applicantRe-Registration/',
    # ------------------------------------------------------------------------------------------------------------------

    "Schedule": "https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",

    "givefeedback": "https://amsin.hirepro.in/py/crpo/api/v1/interview/givefeedback/",

    "Interview_details": "https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}",

    "updateinterviewerdecision": "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerdecision",

    "updateinterviewerfeedback": "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerfeedback",

    "Reschedule": "https://amsin.hirepro.in/py/crpo/api/v1/interview/reschedule/",

    "InterviewRequest_details": "https://amsin.hirepro.in/py/crpo/api/v1/view/interviews",

    "cancel": "https://amsin.hirepro.in/py/crpo/api/v1/interview/cancel/",

    "get_all_candidates": "https://amsin.hirepro.in/py/rpo/get_all_candidates/",

    "gettaskbycandidate": "https://amsin.hirepro.in/py/pofu/api/v1/get-task-by-candidate/",

    "submitform": "https://amsin.hirepro.in/py/pofu/api/v1/submit-form/",

    "Approve_task": "https://amsin.hirepro.in/py/pofu/api/v1/update-candidate-task-status/",

    # ----------
    # POFU APIS
    # ----------
    "bulkimport": "https://amsin.hirepro.in/py/pofu/api/v1/bulkimport",

    # ----------
    # Rpo APIS
    # ----------
    "get_all_jobs": "https://amsin.hirepro.in/py/rpo/get_all_jobs/"

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
        'https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/setApplicantCommunicationStatus',

    "sendVerificationNotification": 'https://amsin.hirepro.in/py/crpo/candidate/api/v1/sendVerificationNotification/',

    "getRegistrationLinkForApplicants":
        'https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/getRegistrationLinkForApplicants/',

    "applicantRe-Registration":
        'https://amsin.hirepro.in/py/crpo/candidatecommunication/api/v1/applicantRe-Registration/',

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
