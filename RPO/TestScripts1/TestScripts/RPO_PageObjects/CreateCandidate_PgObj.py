class createCandidate_PgObj():

    @staticmethod
    def createCandidate_PgElements(driver):
        __elements = {}
        __elements["Upload_Resume"] = driver.find_element_by_css_selector("input[type=\'file\']")
        __elements["Profile_Picture"] = driver.find_element_by_xpath("(//input[@type='file'])[2]")
        __elements["Candidate_Name"] = driver.find_element_by_name("candidatename")
        __elements["Email"] = driver.find_element_by_name("email")
        __elements["Alternate_Email"] = driver.find_element_by_name("alternateEmail")
        __elements["DOB"] = driver.find_element_by_xpath("//div[5]/md-dialog/md-dialog-content/div[2]/div/div/div[2]/form/div/div/div[7]/div[4]/md-datepicker/div/input")
        __elements["Gender"] = driver.find_element_by_name("gender")
        __elements["Location"] = driver.find_element_by_name("location")
        __elements["Mobile"] = driver.find_element_by_name("mobile1")
        __elements["Phone_No"] = driver.find_element_by_name("phoneoffice")
        __elements["Sensitivity"] = driver.find_element_by_name("sensitivity")
        __elements["Candidate_Status"] = driver.find_element_by_name("status")
        __elements["Candidate_Sourcer"] = driver.find_element_by_name("sourcer")
        __elements["Expertise"] = driver.find_element_by_name("expertise")
        __elements["Experience_InYear"] = driver.find_element_by_name("experienceyears")
        __elements["Experience_InMonths"] = driver.find_element_by_name("experiencemonths")
        __elements["Current_Salary_LPA"] = driver.find_element_by_name("currentctc")
        __elements["Source_Type"] = driver.find_element_by_name("sourcetype")
        __elements["Source"] = driver.find_element_by_name("source")
        __elements["Willing_To_Relocate"] = driver.find_element_by_name("willingtorelocate")
        __elements["Expected_Salary_From_LPA"] = driver.find_element_by_name("salaryfrom")
        __elements["Expected_Salary_To_LPA"] = driver.find_element_by_name("salaryto")
        __elements["Notice_Period_Days"] = driver.find_element_by_name("noticeperiod")
        __elements["College"] = driver.find_element_by_name("addedcollege")
        __elements["Degree"] = driver.find_element_by_name("addeddegree")
        __elements["Branch"] = driver.find_element_by_name("addedbranch")
        __elements["YOP"] = driver.find_element_by_name("yop")
        __elements["CGPA/%"] = driver.find_element_by_name("cgpa")
        __elements["IsFinal_Education"] = driver.find_element_by_xpath("//md-input-container/md-checkbox/div")
        __elements["Company"] = driver.find_element_by_name("company")
        __elements["Designation/Role"] = driver.find_element_by_name("designation")
        __elements["Experience_From"] = driver.find_element_by_name("fromyear")
        __elements["Experience_To"] = driver.find_element_by_name("toyear")
        __elements["Salary"] = driver.find_element_by_name("newsalary")
        __elements["Reason_For_Leaving"] = driver.find_element_by_name("addedreasonOfLeaving")
        __elements["IsLatest_Experience"] = driver.find_element_by_xpath("//div[18]/div/div[6]/md-input-container/md-checkbox/div")
        # __elements["Cancel"] =
        # __elements["Save"] =
        return __elements





    # __createCandidate_PgElement = createCandidate_PgObj.createCandidate_PgElements(self.driver)
    # __upload_Resume = __createCandidate_PgElement["Upload_Resume"]
    # __profile_Picture = __createCandidate_PgElement["Profile_Picture"]
    # __candidate_Name = __createCandidate_PgElement["Candidate_Name"]
    # __candidate_Primary_Email = __createCandidate_PgElement["Email"]
    # __candidate_Secondary_Email = __createCandidate_PgElement["Alternate_Email"]
    # __candidate_DOB = __createCandidate_PgElement["DOB"]
    # __gender = __createCandidate_PgElement["Gender"]
    # __location = __createCandidate_PgElement["Location"]
    # __mobile1 = __createCandidate_PgElement["Mobile"]
    # __phone_Number = __createCandidate_PgElement["Phone_No"]
    # __isFinal_Education = __createCandidate_PgElement["IsFinal_Education"]
    # __isFinal_Experience = __createCandidate_PgElement["IsLatest_Experience"]
    # __sensitivity = __createCandidate_PgElement["Sensitivity"]
    # __status = __createCandidate_PgElement["Candidate_Status"]
    # __sourcer = __createCandidate_PgElement["Candidate_Sourcer"]
    # __expertise = __createCandidate_PgElement["Expertise"]
    # __exp_In_Years = __createCandidate_PgElement["Experience_InYear"]
    # __exp_In_Months = __createCandidate_PgElement["Experience_InMonths"]
    # __current_Salary = __createCandidate_PgElement["Current_Salary_LPA"]
    # __source_Type = __createCandidate_PgElement["Source_Type"]
    # __source = __createCandidate_PgElement["Source"]
    # __willing_To_Relocate = __createCandidate_PgElement["Willing_To_Relocate"]
    # __expected_Salary_From = __createCandidate_PgElement["Expected_Salary_From_LPA"]
    # __expected_Salary_To = __createCandidate_PgElement["Expected_Salary_To_LPA"]
    # __notice_Period = __createCandidate_PgElement["Notice_Period_Days"]
    # __college_Name = __createCandidate_PgElement["College"]
    # __degree = __createCandidate_PgElement["Degree"]
    # __branch = __createCandidate_PgElement["Branch"]
    # __yOP = __createCandidate_PgElement["YOP"]
    # __percentage_Or_CGPA = __createCandidate_PgElement["CGPA/%"]
    # __company = __createCandidate_PgElement["Company"]
    # __designation = __createCandidate_PgElement["Designation/Role"]
    # __experience_From = __createCandidate_PgElement["Experience_From"]
    # __experience_To = __createCandidate_PgElement["Experience_To"]
    # __salary = __createCandidate_PgElement["Salary"]
    # __reason_For_Leaving = __createCandidate_PgElement["Reason_For_Leaving"]