import xlrd
import datetime
class Excel_Data:
    def __init__(self):

        self.xl_job_id = []
        self.xl_job_name = []
        self.xl_openings_from = []
        self.xl_openings_to = []
        self.xl_roles = []
        self.xl_users = []
        self.xl_location = []
        self.xl_hiring_type = []
        self.xl_utilized = []
        self.xl_department = []

        self.xl_sub_department = []
        self.xl_job_code = []
        self.xl_ctc_from = []
        self.xl_ctc_to = []


        self.xl_experience_from = []
        self.xl_experience_to = []
        self.xl_is_posted = []
        self.xl_job_skills = []



    def Data_read(self):
        wb = xlrd.open_workbook('/home/sanjeev/TestScripts/TestScripts/Job_Search.xls')
        sheetname = wb.sheet_names()  # Reading XLS Sheet names
        # print(sheetname)
        sh1 = wb.sheet_by_index(0)  #
        i = 1
        for i in range (1,sh1.nrows):
            rownum = (i)
            rows = sh1.row_values(rownum)
            self.xl_job_id.append(int(rows[0]))
            self.xl_job_name.append(rows[1])
            self.xl_openings_from.append(int(rows[2]))
            self.xl_openings_to.append(int(rows[3]))
            self.xl_roles.append(int(rows[4]))
            self.xl_users.append(int(rows[5]))
            self.xl_location.append(int(rows[6]))
            self.xl_hiring_type.append(int(rows[7]))
            self.xl_utilized.append(int(rows[8]))
            self.xl_department.append(int(rows[9]))

            self.xl_sub_department.append(int(rows[10]))
            self.xl_job_code.append(int(rows[11]))
            self.xl_ctc_from.append(int(rows[12]))
            self.xl_ctc_to.append(int(rows[13]))
            self.xl_experience_from.append(int(rows[14]))
            self.xl_experience_to.append(int(rows[15]))
            self.xl_is_posted.append(int(rows[16]))
            self.xl_job_skills.append(int(rows[17]))




    def disp(self):
        print self.xl_job_id
        print  self.xl_job_name
        print self.xl_openings_from
        print self.xl_openings_to
        print  self.xl_roles
        print self.xl_users
        print self.xl_location
        print  self.xl_hiring_type
        print self.xl_experience_from
        print self.xl_experience_to
        print  self.xl_utilized
        print self.xl_department
        print self.xl_sub_department
        print  self.xl_job_code
        print self.xl_ctc_from
        print self.xl_ctc_to
        print  self.xl_experience_from
        print self.xl_experience_to
        print self.xl_is_posted
        print  self.xl_job_skills




xlob = Excel_Data()
xlob.Data_read()
# xlob.disp()
# print xlob.xl_created_on_from
# print xlob.xl_created_on_to