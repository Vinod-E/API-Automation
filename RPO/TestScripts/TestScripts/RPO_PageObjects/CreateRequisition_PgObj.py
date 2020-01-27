class createRequisition_PgObj():

    @staticmethod
    def createRequisition_PgElements(driver):
        __elements = {}
        __elements["Customer"] = driver.find_element_by_name("company")
        __elements["Job_Code"] = driver.find_element_by_name("reqcode")
        __elements["Req_Title"] = driver.find_element_by_name("reqtitle")
        __elements["Openings"] = driver.find_element_by_name("openings")
        __elements["Location"] = driver.find_element_by_name("location")
        __elements["Req_Type"] = driver.find_element_by_name("reqtype")
        __elements["Experience_Range"] = driver.find_element_by_name("experiencerangefrom")
        __elements["Salary_From_LPA"] = driver.find_element_by_name("salaryfrom")
        __elements["Salary_To_LPA"] = driver.find_element_by_name("salaryto")
        __elements["Designation"] = driver.find_element_by_name("grade")
        __elements["Expertise"] = driver.find_element_by_name("expertise")
        __elements["Role"] = driver.find_element_by_name("role")
        __elements["Technology_Text"] = driver.find_element_by_xpath("//div[8]/md-input-container/md-select")
        __elements["Sensitivity"] = driver.find_element_by_name("sensitivity")
        __elements["Requisition_Owner"] = driver.find_element_by_name("reqowner")
        __elements["Requisition_Approver"] = driver.find_element_by_name("reqapprover")
        __elements["Recruiter"] = driver.find_element_by_name("customertaowner")
        __elements["Requisition_Name"] = driver.find_element_by_name("reqName")
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