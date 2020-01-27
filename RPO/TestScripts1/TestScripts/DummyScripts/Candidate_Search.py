import mysql.connector
import requests
import json
from Job_Search_Input import *
import datetime
import xlwt
from collections import OrderedDict

class AMS_DB_Data:
    def __init__(self):
        self.conn = mysql.connector.connect(host='35.154.36.218',
                                            database='core1779',
                                            user='hireprouser',
                                            password='tech@123')
        self.cursor = self.conn.cursor()

    def execute_Query(self, query):
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchone()
        except:
            print("Hi")

    def ams_Query(self,b):

        self.candidate_name = "select count(1) from candidates where candidate_name like '%s' and is_archived=0 " \
                              "and is_deleted=0 and is_draft=0;"% ('%' + xlob.xl_candidate_name[b] + '%')
        # print xlob.xl_candidate_name[b]
        query = self.candidate_name
        # print query
        C4.execute_Query(query)
        self.db_candidate_name = self.data[0]
        # print "total candidate name count is %s"%self.db_candidate_name

        self.candidate_email = "select count(id) from candidates where email1 like '%s' or " \
                               "email2  like '%s' and is_archived=0 " \
                               "and is_deleted=0 and is_draft=0;" %('%'+xlob.xl_email[b] + '%','%'+ xlob.xl_email[b] +'%')
        query = self.candidate_email
        # print query
        C4.execute_Query(query)
        self.db_candidate_email = self.data[0]
        # print "total candidate_emailcount is %s" % self.db_candidate_email

        self.candidate_mobile = "select count(1) from candidates where mobile1 like '%s' and is_archived=0 " \
                                "and is_deleted=0 and is_draft=0;"%(xlob.xl_mobile[b])
        query = self.candidate_mobile
        # print query
        C4.execute_Query(query)
        self.db_candidate_mobile = self.data[0]
        # print "total candidate_mobile count is %s" % self.db_candidate_mobile

        self.candidate_job_department = "select count(distinct(a.candidate_id)) from applicant_statuss a left join jobs " \
                                        "j on a.job_id = j.id left join candidates c on a.candidate_id = c.id " \
                                        "where a.job_id in (select id from jobs where department_id='%s') " \
                                        "and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_jobrole_department_id[b])
        query = self.candidate_job_department
        # print query
        C4.execute_Query(query)
        self.db_candidate_job_department = self.data[0]
        # print "total candidate_mobile count is %s" % self.db_candidate_mobile


# Get Name
        self.db_job_role_department_name = "select department_name from departments where id=%s" %(xlob.xl_jobrole_department_id[b])
        query = self.db_job_role_department_name
        # print query
        C4.execute_Query(query)
        self.db_job_role_department_name = self.data[0]
        # print "Department Name is %s" % self.db_job_role_department_name


        self.candidate_jobapplicant = "select count(c.id) from applicant_statuss a left join candidates c on c.id=a.candidate_id " \
                                "where job_id='%s' and c.is_archived=0 and c.is_deleted=0 and " \
                                      "c.is_draft=0 and c.tenant_id=1779;" %(xlob.xl_jobrole_id[b])
        query = self.candidate_jobapplicant
        # print query
        C4.execute_Query(query)
        self.db_candidate_jobapplicant = self.data[0]
        # print "total candidate_jobapplicant count is %s" % self.db_candidate_jobapplicant

# Get Name
        self.db_job_name = "select job_name from jobs where id=%s" %(xlob.xl_jobrole_id[b])
        query = self.db_job_name
        # print query
        C4.execute_Query(query)
        self.db_job_name = self.data[0]
        # print "Job Name is %s" % self.db_job_name




        self.candidate_stage_status = "select count(distinct(a.candidate_id)) from applicant_statuss a left join candidates c " \
                                      "on c.id=a.candidate_id where a.current_status_id='%s' and c.is_archived=0 and " \
                                      "c.is_deleted=0 and c.is_draft=0 and a.is_deleted=0 and " \
                                      "c.tenant_id=1779;"%(xlob.xl_stage_status_id[b])
        query = self.candidate_stage_status
        # print query
        C4.execute_Query(query)
        self.db_candidate_stage_status = self.data[0]
        # print "total candidate_stage_status count is %s" % self.db_candidate_stage_status
# Get Name
        self.db_stage_status = "select rs2.id,rs1.label,rs2.base_name from resume_statuss rs1,resume_statuss rs2 " \
                       "where  rs1.id = rs2.resumestatus_id and rs2.id = % s;" %(xlob.xl_stage_status_id[b])
        query = self.db_stage_status
        C4.execute_Query(query)
        # print query
        self.db_stage_name = self.data[1]
        self.db_status_name = self.data[2]
        # self.db_stage_name = "CV"
        # self.db_status_name = "Matching"
        self.db_stage_status_name = "%s - %s"%(self.db_stage_name,self.db_status_name)
        # print "Stage Status Name is %s" % self.db_stage_status


        self.candidate_current_location = "select count(1) from candidates where current_location_id = '%s' and is_archived=0 " \
                                          "and is_deleted=0 and is_draft=0;"%(xlob.xl_current_Location_id[b])
        query = self.candidate_current_location
        # print query
        C4.execute_Query(query)
        self.db_candidate_current_location = self.data[0]
        # print "total candidate_current_location count is %s" % self.db_candidate_current_location

# Get Name
        self.db_location_name = "select value from catalog_values where id=%s" %(xlob.xl_current_Location_id[b])
        query = self.db_location_name
        # print query
        C4.execute_Query(query)
        self.db_location_name= self.data[0]
        # print "Location Name is %s" % self.db_location_name

        self.candidate_total_experience = "select count(1) from candidates where total_experience between %s*12 and %s*12 " \
                                          "and is_archived=0 and is_deleted=0 and is_draft=0;"\
                                          %(xlob.xl_experience_from[b],xlob.xl_experience_to[b])
        query = self.candidate_total_experience
        # print query
        C4.execute_Query(query)
        self.db_candidate_total_experience = self.data[0]
        # print "total candidate_total_experience count is %s" % self.db_candidate_total_experience

        self.candidate_current_employer = "select count(1) from candidates where current_employer_id = '%s' and is_archived=0 " \
                                          "and is_deleted=0 and is_draft=0;"%(xlob.xl_organization_id[b])
        query = self.candidate_current_employer
        # print query
        C4.execute_Query(query)
        self.db_candidate_current_employer = self.data[0]
        # print "total candidate_current_employer count is %s" % self.db_candidate_current_employer

# Get Name
        self.db_employer_name = "select value from catalog_values where id=%s" %(xlob.xl_organization_id[b])
        query = self.db_employer_name
        # print query
        C4.execute_Query(query)
        self.db_employer_name  = self.data[0]
        # print "Employer Name is %s" % self.db_employer_name

        self.candidate_expertise = "select count(1) from candidates where expertise_id1 = '%s' and is_archived=0 " \
                                   "and is_deleted=0 and is_draft=0;"%(xlob.xl_expertise_id[b])
        query = self.candidate_expertise
        # print query
        C4.execute_Query(query)
        self.db_candidate_expertise = self.data[0]
        # print "total candidate_expertise count is %s" % self.db_candidate_expertise
# Get Name
        self.db_expertise_name = "select value from catalog_values where id=%s" %(xlob.xl_expertise_id[b])
        query = self.db_expertise_name
        # print query
        C4.execute_Query(query)
        self.db_expertise_name  = self.data[0]
        # print "Expertise Name is %s" % self.db_expertise_name


        self.candidate_types_of_source = "select count(c.id) from sources s left join candidates " \
                                         "c on s.id = c.original_source_id where c.original_source_id in " \
                                         "(select id from sources where s.types_of_source='%s') and c.is_archived=0 " \
                                         "and c.is_deleted=0 and c.is_draft=0 ;"%(xlob.xl_type_of_source_id[b])
        query = self.candidate_types_of_source
        # print query
        C4.execute_Query(query)
        self.db_types_of_source = self.data[0]
        # print "total candidate_expertise count is %s" % self.db_types_of_source

        self.source = "select count(1) from candidates where original_source_id = '%s' and is_archived=0 " \
                      "and is_deleted=0 and is_draft=0;"%(xlob.xl_source_id[b])
        query = self.source
        # print query
        C4.execute_Query(query)
        self.db_source = self.data[0]
        # print "total source count is %s" % self.db_source
# Get Name
        self.db_source_name = "select source_name from sources where id =%s" %(xlob.xl_source_id[b])
        query = self.db_source_name
        # print query
        C4.execute_Query(query)
        self.db_source_name  = self.data[0]
        # print "Source Name is %s" % self.db_source_name

        self.college = "select count(c.id) from candidates c left join candidate_education_profiles ce " \
                       "on c.id=ce.candidate_id where  ce.college_id = '%s' and c.is_archived=0 and " \
                       "c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_college_id[b])
        query = self.college
        # print query
        C4.execute_Query(query)
        self.db_college = self.data[0]
        # print "total college count is %s" % self.db_college
# Get Name
        self.db_college_name = "select value from catalog_values where id=%s" %(xlob.xl_college_id[b])
        query = self.db_college_name
        # print query
        C4.execute_Query(query)
        self.db_college_name  = self.data[0]
        # print "College Name is %s" % self.db_college_name



        self.degree = "select count(c.id) from candidates c left join candidate_education_profiles ce " \
                      "on c.id=ce.candidate_id where  ce.degree_id = '%s' and c.is_archived=0 " \
                      "and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_degree_id[b])
        query = self.degree
        # print query
        C4.execute_Query(query)
        self.db_degree = self.data[0]
        # print "total degree count is %s" % self.db_degree
# Get Name
        self.db_degree_name = "select value from catalog_values where id=%s" %(xlob.xl_degree_id[b])
        query = self.db_degree_name
        # print query
        C4.execute_Query(query)
        self.db_degree_name  = self.data[0]
        # print "Degree Name is %s" % self.db_degree_name

        self.department = "select count(distinct(c.id)) from candidate_education_profiles cw left join candidates " \
                          "c on c.id= cw.candidate_id where cw.degree_type_id='%s' and c.is_deleted=0 and c.is_archived=0 " \
                          "and c.is_draft=0;"%(xlob.xl_department_id[b])
        query = self.department
        # print query
        C4.execute_Query(query)
        self.db_department = self.data[0]
        # print "total department count is %s" % self.db_department
# Get Name
        self.db_department_name = "select value from catalog_values where id=%s" %(xlob.xl_department_id[b])
        query = self.db_department_name
        # print query
        C4.execute_Query(query)
        self.db_department_name  = self.data[0]
        # print "Department Name is %s" % self.db_department_name

        self.yop = "select count(distinct(c.id)) from candidates c left join candidate_education_profiles ce " \
                   "on c.id=ce.candidate_id where  ce.end_year = '%s' and c.is_archived=0 and c.is_deleted=0 and " \
                   "c.is_draft=0;"%(xlob.xl_yop_id[b])
        query = self.yop
        # print query
        C4.execute_Query(query)
        self.db_yop = self.data[0]
        # print "total yop count is %s" % self.db_yop
# Get Name
        self.db_yop_name = "select value from catalog_values where id=%s" %(xlob.xl_yop_id[b])
        query = self.db_yop_name
        # print query
        C4.execute_Query(query)
        self.db_yop_name  = self.data[0]
        # print "YOP Value is %s" % self.db_yop_name

        self.percentage = "select count(distinct(c.id)) from candidates c left join candidate_education_profiles ce " \
                          "on c.id=ce.candidate_id where  ce.percentage  between '%s' and '%s' and c.is_archived=0 " \
                          "and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_percentage_from[b],xlob.xl_percentage_to[b])
        query = self.percentage
        # print query
        C4.execute_Query(query)
        self.db_percentage = self.data[0]
        # print "total percentage count is %s" % self.db_percentage

        self.marital = "select count(1) from candidates where marital_status = '%s' and is_archived=0 and" \
                       " is_deleted=0 and is_draft=0;"%(xlob.xl_marital_status_id[b])
        query = self.marital
        # print query
        C4.execute_Query(query)
        self.db_marital = self.data[0]
        # print "total marital count is %s" % self.db_marital

        self.sourcer = "select count(id) from candidates where sourcer = '%s' and is_archived=0 and " \
                       "is_deleted=0 and is_draft=0;"%(xlob.xl_sourcer_id[b])
        query = self.sourcer
        # print query
        C4.execute_Query(query)
        self.db_sourcer = self.data[0]
        # print "total sourcer count is %s" % self.db_sourcer
# Get Name
        self.db_sourcer_name = "select user_name from users where id=%s" %(xlob.xl_sourcer_id[b])
        query = self.db_sourcer_name
        # print query
        C4.execute_Query(query)
        self.db_sourcer_name  = self.data[0]
        # print "Sourcer Name is %s" % self.db_sourcer_name

        self.created_on = "select count(id) from candidates where date(created_on) between '%s' and '%s' " \
                          "and is_archived=0 and is_deleted=0 and is_draft=0;"%(xlob.xl_created_on_from[b],xlob.xl_created_on_to[b])
        query = self.created_on
        # print query
        C4.execute_Query(query)
        self.db_created_on = self.data[0]
        # print "total created_on count is %s" % self.db_created_on

        self.modefied_on = "select count(id) from candidates where date(modified_on) between '%s' and '%s' " \
                           "and is_archived=0 and is_deleted=0 and is_draft=0;"%(xlob.xl_modified_on_from[b],xlob.xl_modified_on_to[b])
        query = self.modefied_on
        # print query
        C4.execute_Query(query)
        self.db_modefied_on = self.data[0]
        # print "total modefied_on count is %s" % self.db_modefied_on

        self.created_by = "select count(1) from candidates where created_by= '%s' and is_archived=0 and is_deleted=0 " \
                          "and is_draft=0;"%(xlob.xl_created_by_id[b])
        query = self.created_by
        # print query
        C4.execute_Query(query)
        self.db_created_by = self.data[0]
        # print "total created_by count is %s" % self.db_created_by
# Get Name
        self.db_createdby_name = "select user_name from users where id=%s" %(xlob.xl_created_by_id[b])
        query = self.db_createdby_name
        # print query
        C4.execute_Query(query)
        self.db_createdby_name  = self.data[0]
        # print "Created By Name is %s" % self.db_createdby_name

        self.is_utilized = "select count(1) from candidates where is_utilized= 0 and is_archived=0 and is_deleted=0 and is_draft=0;"
        query = self.is_utilized
        # print query
        C4.execute_Query(query)
        self.db_is_utilized = self.data[0]
        # print "total is_utilized count is %s" % self.db_is_utilized

        self.is_gender = "select count(1) from candidates where gender= '%s' and is_archived=0 and is_deleted=0 " \
                         "and is_draft=0;"%(xlob.xl_is_gender_id[b])
        query = self.is_gender
        # print query
        C4.execute_Query(query)
        self.db_is_gender = self.data[0]
        # print "total is_gender count is %s" % self.db_is_gender

        self.is_usn = "select count(1) from candidates where usn= '%s' and is_archived=0 and is_deleted=0 " \
                      "and is_draft=0;"%(xlob.xl_usn[b])
        query = self.is_usn
        # print query
        C4.execute_Query(query)
        self.db_is_usn = self.data[0]
        # print "total is_usn count is %s" % self.db_is_usn


        self.text1 = "select count(1) from candidates where text1 like '%s' and is_archived=0 and " \
                     "is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_text1[b] + '%')
        query = self.text1
        # print query
        C4.execute_Query(query)
        self.db_text1 = self.data[0]
        # print "total text1 count is %s" % self.db_text1

        self.text2 = "select count(1) from candidates where text2 like '%s' and is_archived=0 and " \
                     "is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_text2[b] + '%')
        query = self.text2
        # print query
        C4.execute_Query(query)
        self.db_text2 = self.data[0]
        # print "total text2 count is %s" % self.db_text2

        self.text3 = "select count(1) from candidates where text3 like '%s' and is_archived=0 and " \
                     "is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_text3[b] + '%')
        query = self.text3
        # print query
        C4.execute_Query(query)
        self.db_text3 = self.data[0]
        # print "total text3 count is %s" % self.db_text3

        self.text4 = "select count(1) from candidates where text4 like '%s' and is_archived=0 and " \
                     "is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_text4[b] + '%')
        query = self.text4
        # print query
        C4.execute_Query(query)
        self.db_text4= self.data[0]
        # print "total text4 count is %s" % self.db_text4

        self.text5 = "select count(1) from candidates where text5 like '%s' and is_archived=0 and " \
                     "is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_text5[b] + '%')
        query = self.text5
        # print query
        C4.execute_Query(query)
        self.db_text5 = self.data[0]
        # print "total text5 count is %s" % self.db_text5

        self.text6 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text6 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text6[b] + '%')
        query = self.text6
        # print query
        C4.execute_Query(query)
        self.db_text6 = self.data[0]
        # print "total text6 count is %s" % self.db_text6

        self.text7 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text7 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text7[b] + '%')
        query = self.text7
        # print query
        C4.execute_Query(query)
        self.db_text7 = self.data[0]
        # print "total text7 count is %s" % self.db_text7

        self.text8 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text8 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text8[b] + '%')
        C4.execute_Query(query)
        # print query
        self.db_text8 = self.data[0]
        # print "total text8 count is %s" % self.db_text8

        self.text9 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text9 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text9[b] + '%')
        query = self.text9
        # print query
        C4.execute_Query(query)
        self.db_text9 = self.data[0]
        # print "total text9 count is %s" % self.db_text9

        self.text10 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text10 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text10[b] + '%')
        query = self.text10
        # print query
        C4.execute_Query(query)
        self.db_text10 = self.data[0]
        # print "total text10 count is %s" % self.db_text10

        self.text11 ="select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text11 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text11[b] + '%')
        query = self.text11
        # print query
        C4.execute_Query(query)
        self.db_text11 = self.data[0]
        # print "total text11 count is %s" % self.db_text11

        self.text12 ="select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text12 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text12[b] + '%')
        query = self.text12
        # print query
        C4.execute_Query(query)
        self.db_text12 = self.data[0]
        # print "total text12 count is %s" % self.db_text12

        self.text13 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text13 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text13[b] + '%')
        query = self.text13
        # print query
        C4.execute_Query(query)
        self.db_text13 = self.data[0]
        # print "total text13 count is %s" % self.db_text13

        self.text14 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text14 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text14[b] + '%')
        query = self.text14
        # print query
        C4.execute_Query(query)
        self.db_text14 = self.data[0]
        # print "total text14 count is %s" % self.db_text14

        self.text15 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where cust.text15 like '%s' and c.is_archived=0 and c.is_deleted=0 " \
                     "and c.is_draft=0;"%('%'+ xlob.xl_text15[b] + '%')
        query = self.text15
        # print query
        C4.execute_Query(query)
        self.db_text15 = self.data[0]
        # print "total text15 count is %s" % self.db_text15

        self.integer1 = "select count(1) from candidates where integer1 = '%s' and is_archived=0 and " \
                        "is_deleted=0 and is_draft=0;"%(xlob.xl_integer1[b])
        query = self.integer1
        # print query
        C4.execute_Query(query)
        self.db_integer1 = self.data[0]
        # print "total integer1 count is %s" % self.db_integer1
# Get Name
        self.db_integer1_name = "select value from catalog_values where id=%s" %(xlob.xl_integer1[b])
        query = self.db_integer1_name
        # print query
        C4.execute_Query(query)
        self.db_integer1_name  = self.data[0]
        # print "Integer1 Value is %s" % self.db_integer1_name

        self.integer2 = "select count(1) from candidates where integer2 = '%s' and is_archived=0 and " \
                        "is_deleted=0 and is_draft=0;"%(xlob.xl_integer2[b])
        query = self.integer2
        # print query
        C4.execute_Query(query)
        self.db_integer2 = self.data[0]
        # print "total integer2 count is %s" % self.db_integer2
# Get Name
        self.db_integer2_name = "select value from catalog_values where id=%s" %(xlob.xl_integer2[b])
        query = self.db_integer2_name
        # print query
        C4.execute_Query(query)
        self.db_integer2_name  = self.data[0]
        # print "Integer2 Value is %s" % self.db_integer2_name


        self.integer3 =  "select count(1) from candidates where integer3 = '%s' and is_archived=0 and " \
                        "is_deleted=0 and is_draft=0;"%(xlob.xl_integer3[b])
        query = self.integer3
        # print query
        C4.execute_Query(query)
        self.db_integer3 = self.data[0]
        # print "total integer3 count is %s" % self.db_integer3
# Get Name
        self.db_integer3_name = "select value from catalog_values where id=%s" %(xlob.xl_integer3[b])
        query = self.db_integer3_name
        # print query
        C4.execute_Query(query)
        self.db_integer3_name  = self.data[0]
        # print "Integer3 Value is %s" % self.db_integer3_name


        self.integer4 =  "select count(1) from candidates where integer4 = '%s' and is_archived=0 and " \
                        "is_deleted=0 and is_draft=0;"%(xlob.xl_integer4[b])
        query = self.integer4
        # print query
        C4.execute_Query(query)
        self.db_integer4 = self.data[0]
        # print "total integer4 count is %s" % self.db_integer4
# Get Name
        self.db_integer4_name = "select value from catalog_values where id=%s" %(xlob.xl_integer4[b])
        query = self.db_integer4_name
        # print query
        C4.execute_Query(query)
        self.db_integer4_name  = self.data[0]
        # print "Integer4 Value is %s" % self.db_integer4_name


        self.integer5 =  "select count(1) from candidates where integer5 = '%s' and is_archived=0 and " \
                        "is_deleted=0 and is_draft=0;"%(xlob.xl_integer5[b])
        query = self.integer5
        # print query
        C4.execute_Query(query)
        self.db_integer5 = self.data[0]
        # print "total integer5 count is %s" % self.db_integer5
# Get Name
        self.db_integer5_name = "select value from catalog_values where id=%s" %(xlob.xl_integer5[b])
        query = self.db_integer5_name
        # print query
        C4.execute_Query(query)
        self.db_integer5_name  = self.data[0]
        # print "Integer5 Value is %s" % self.db_integer5_name


        self.integer6 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer6 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer6[b])
        query = self.integer6
        # print query
        C4.execute_Query(query)
        self.db_integer6 = self.data[0]
        # print "total integer6 count is %s" % self.db_integer6
# Get Name
        self.db_integer6_name = "select value from catalog_values where id=%s" %(xlob.xl_integer6[b])
        query = self.db_integer6_name
        # print query
        C4.execute_Query(query)
        self.db_integer6_name  = self.data[0]
        # print "Integer6 Value is %s" % self.db_integer6_name


        self.integer7 =  "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer7 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer7[b])
        query = self.integer7
        # print query
        C4.execute_Query(query)
        self.db_integer7 = self.data[0]
        # print "total integer7 count is %s" % self.db_integer7
# Get Name
        self.db_integer7_name = "select value from catalog_values where id=%s" %(xlob.xl_integer7[b])
        query = self.db_integer7_name
        # print query
        C4.execute_Query(query)
        self.db_integer7_name  = self.data[0]
        # print "Integer7 Value is %s" % self.db_integer7_name


        self.integer8 =  "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer8 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer8[b])
        query = self.integer8
        # print query
        C4.execute_Query(query)
        self.db_integer8 = self.data[0]
        # print "total integer8 count is %s" % self.db_integer8
# Get Name
        self.db_integer8_name = "select value from catalog_values where id=%s" %(xlob.xl_integer8[b])
        query = self.db_integer8_name
        # print query
        C4.execute_Query(query)
        self.db_integer8_name  = self.data[0]
        # print "Integer8 Value is %s" % self.db_integer8_name


        self.integer9 =  "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer9 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer9[b])
        query = self.integer9
        # print query
        C4.execute_Query(query)
        self.db_integer9 = self.data[0]
        # print "total integer9 count is %s" % self.db_integer9
# Get Name
        self.db_integer9_name = "select value from catalog_values where id=%s" %(xlob.xl_integer9[b])
        query = self.db_integer9_name
        # print query
        C4.execute_Query(query)
        self.db_integer9_name  = self.data[0]
        # print "Integer9 Value is %s" % self.db_integer9_name


        self.integer10 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer10 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer10[b])
        query = self.integer10
        # print query
        C4.execute_Query(query)
        self.db_integer10 = self.data[0]
        # print "total integer10 count is %s" % self.db_integer10
# Get Name
        self.db_integer10_name = "select value from catalog_values where id=%s" %(xlob.xl_integer10[b])
        query = self.db_integer10_name
        # print query
        C4.execute_Query(query)
        self.db_integer10_name  = self.data[0]
        # print "Integer10 Value is %s" % self.db_integer10_name


        self.integer11 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer11 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer11[b])
        query = self.integer11
        # print query
        C4.execute_Query(query)
        self.db_integer11 = self.data[0]
        # print "total integer11 count is %s" % self.db_integer11
# Get Name
        self.db_integer11_name = "select value from catalog_values where id=%s" %(xlob.xl_integer11[b])
        query = self.db_integer11_name
        # print query
        C4.execute_Query(query)
        self.db_integer11_name  = self.data[0]
        # print "Integer11 Value is %s" % self.db_integer11_name

        self.integer12 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer12 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer12[b])
        query = self.integer12
        # print query
        C4.execute_Query(query)
        self.db_integer12 = self.data[0]
        # print "total integer12 count is %s" % self.db_integer12
# Get Name
        self.db_integer12_name = "select value from catalog_values where id=%s" %(xlob.xl_integer12[b])
        query = self.db_integer12_name
        # print query
        C4.execute_Query(query)
        self.db_integer12_name  = self.data[0]
        # print "Integer12 Value is %s" % self.db_integer12_name

        self.integer13 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer13 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer13[b])
        query = self.integer13
        # print query
        C4.execute_Query(query)
        self.db_integer13 = self.data[0]
        # print "total integer13 count is %s" % self.db_integer13
# Get Name
        self.db_integer13_name = "select value from catalog_values where id=%s" %(xlob.xl_integer13[b])
        query = self.db_integer13_name
        # print query
        C4.execute_Query(query)
        self.db_integer13_name  = self.data[0]
        # print "Integer13 Value is %s" % self.db_integer13_name

        self.integer14 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer14 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer14[b])
        query = self.integer14
        # print query
        C4.execute_Query(query)
        self.db_integer14 = self.data[0]
        # print "total integer14 count is %s" % self.db_integer14

# Get Name
        self.db_integer14_name = "select value from catalog_values where id=%s" %(xlob.xl_integer14[b])
        query = self.db_integer14_name
        # print query
        C4.execute_Query(query)
        self.db_integer14_name  = self.data[0]
        # print "Integer14 Value is %s" % self.db_integer14_name

        self.integer15 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                        "where cust.integer15 = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_integer15[b])
        query = self.integer15
        # print query
        C4.execute_Query(query)
        self.db_integer15 = self.data[0]
        # print "total integer15 count is %s" % self.db_integer15
# Get Name
        self.db_integer15_name = "select value from catalog_values where id=%s" %(xlob.xl_integer15[b])
        query = self.db_integer15_name
        # print query
        C4.execute_Query(query)
        self.db_integer15_name  = self.data[0]
        # print "Integer15 Value is %s" % self.db_integer15_name


        self.textarea1 = "select count(1) from candidates where text_area1 like '%s' and is_archived=0" \
                         " and is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_textarea1[b] +'%')
        query = self.textarea1
        # print query
        C4.execute_Query(query)
        self.db_textarea1 = self.data[0]
        # print "total textarea1 count is %s" % self.db_textarea1

        self.textarea2 =  "select count(1) from candidates where text_area2 like '%s' and is_archived=0" \
                         " and is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_textarea2[b] +'%')
        query = self.textarea2
        # print query
        C4.execute_Query(query)
        self.db_textarea2 = self.data[0]
        # print "total textarea2 count is %s" % self.db_textarea2

        self.textarea3 =  "select count(1) from candidates where text_area3 like '%s' and is_archived=0" \
                         " and is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_textarea3[b] +'%')
        query = self.textarea3
        # print query
        C4.execute_Query(query)
        self.db_textarea3 = self.data[0]
        # print "total textarea3 count is %s" % self.db_textarea3

        self.textarea4 =  "select count(1) from candidates where text_area4 like '%s' and is_archived=0" \
                         " and is_deleted=0 and is_draft=0;"%('%'+ xlob.xl_textarea4[b] +'%')
        query = self.textarea4
        # print query
        C4.execute_Query(query)
        self.db_textarea4 = self.data[0]
        # print "total textarea4 count is %s" % self.db_textarea4

        self.date1 = "select count(id) from candidates where date(date_custom_field1) = '%s' " \
                     "and is_archived=0 and is_deleted=0 and is_draft=0;"%(xlob.xl_date1[b])
        query = self.date1
        # print query
        C4.execute_Query(query)
        self.db_date1 = self.data[0]
        # print "total date1 count is %s" % self.db_date1

        self.date2 = "select count(id) from candidates where date(date_custom_field2) = '%s' " \
                     "and is_archived=0 and is_deleted=0 and is_draft=0;"%(xlob.xl_date2[b])
        query = self.date2
        print query
        C4.execute_Query(query)
        self.db_date2 = self.data[0]
        # print "total date2 count is %s" % self.db_date2

        self.date3 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where date(cust.date_custom_field3) = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_date3[b])
        query = self.date3
        # print query
        C4.execute_Query(query)
        self.db_date3 = self.data[0]
        # print "total date3 count is %s" % self.db_date3

        self.date4 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where date(cust.date_custom_field4) = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_date4[b])
        query = self.date4
        # print query
        C4.execute_Query(query)
        self.db_date4 = self.data[0]
        # print "total date4 count is %s" % self.db_date4

        self.date5 = "select count(1) from candidates c left join candidate_customs cust on c.candidatecustom_id=cust.id " \
                     "where date(cust.date_custom_field5) = '%s' and c.is_archived=0 and c.is_deleted=0 and c.is_draft=0;"%(xlob.xl_date1[b])
        query = self.date5
        # print query
        C4.execute_Query(query)
        self.db_date5 = self.data[0]
        # print "total date4 count is %s" % self.db_date5

        # self.conn.close()

class api_data:

    def json_data(self,):

        r = requests.post("https://amsin.hirepro.in/py/rpo/get_all_candidates/", headers = self.headers,
                          data = json.dumps(self.data, default=str), verify=False)
        # print r.content
        resp_dict = json.loads(r.content)
        self.status = resp_dict['status']
        # print self.status
        # print type(self.status)
        if self.status == 'OK':
            self.count = resp_dict['TotalItem']
            # print self.count
        else:
            self.count = "400000000000000"
            # print self.count


    def api_main(self,b):
        headers1 = {"content-type": "application/json"}
        data = {"LoginName":"admin","Password":"admin@123","TenantAlias":"rpotest","UserName":"admin"}
        response = requests.post("https://amsin.hirepro.in/py/common/user/login_user/", headers=headers1,
                                 data=json.dumps(data), verify=False)
        abc = response.json()
        # print abc.get("Token")
        # print json.loads(response)
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": abc.get("Token")}

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Name": ""}, "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Name"] = xlob.xl_candidate_name[b]
        # print self.data
        ob.json_data()
        self.api_candidate_name = self.count
        # print "Candidate name Count is %s " %self.api_candidate_name

        # Candidate Email Count
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"Email":"rpmuthumurugan@gmail.com"},"IsNotCacheRequired":False}
        self.data ["CandidateFilters"]["Email"] = xlob.xl_email[b]
        ob.json_data()
        self.api_candidate_email = self.count
        # print "api_candidate_email Count is %s " %self.api_candidate_email

        # Candidate Mobile Count
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"Phone":""},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["Phone"] = xlob.xl_mobile[b]
        ob.json_data()
        self.api_candidate_mobile = self.count
        # print "api_candidate_mobile Count is %s " %self.api_candidate_mobile

        # job_department candidates Count
        self.data ={"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                    "ApplicantFilters":{"JobDepartmentIds":[""]},"IsNotCacheRequired":False}
        self.data["ApplicantFilters"]["JobDepartmentIds"] = [int(xlob.xl_jobrole_department_id[b])]
        # print self.data
        ob.json_data()
        self.api_candidate_job_department = self.count
        # print "api_candidate_job_department Count is %s " %self.api_candidate_job_department

        # Candidate Jobrole Count
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "ApplicantFilters":{"JobIds":[""]},"IsNotCacheRequired":False}
        self.data["ApplicantFilters"]["JobIds"] = [int(xlob.xl_jobrole_id[b])]
        # print self.data
        ob.json_data()
        self.api_jobrole_applicants = self.count
        # print "api_jobrole_applicants is %s " %self.api_jobrole_applicants

        # Candidate Stage Status Count
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "ApplicantFilters":{"ApplicantStatusIds":[""]},"IsNotCacheRequired":False}
        self.data["ApplicantFilters"]["ApplicantStatusIds"] = [int(xlob.xl_stage_status_id[b])]
        ob.json_data()
        self.api_candidate_stage_status = self.count
        # print "api_candidate_stage_status Count is %s " %self.api_candidate_stage_status

        # Candidate Current Location
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"CurrentLocationIds":["25151"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["CurrentLocationIds"] = [int(xlob.xl_current_Location_id[b])]
        # print self.data
        ob.json_data()
        self.api_candidate_current_location = self.count
        # print "api_candidate_current_location Count is %s " %self.api_candidate_current_location

        # Candidate Experience
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"ExperienceFrom":3,"ExperienceTo":5},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["ExperienceFrom"] = int(xlob.xl_experience_from[b])
        self.data["CandidateFilters"]["ExperienceTo"] = int(xlob.xl_experience_to[b])
        # print self.data
        ob.json_data()
        self.api_candidate_experience = self.count
        # print "api_candidate_experience Count is %s " %self.api_candidate_experience

        # Candidate Organization
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"WorkProfileOrganisationIds":["325"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["WorkProfileOrganisationIds"] = [int(xlob.xl_organization_id[b])]
        # print self.data
        ob.json_data()
        self.api_candidate_organization = self.count
        # print "api_candidate_organization Count is %s " %self.api_candidate_organization

        # Candidate Expertise
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"ExpertiseIds":["1627"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["ExpertiseIds"] = [int(xlob.xl_expertise_id[b])]
        # print self.data
        ob.json_data()
        self.api_candidate_expertise = self.count
        # print "api_candidate_expertise Count is %s " %self.api_candidate_expertise


        # Candidate Types Of Source
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"TypesOfSource":"1"},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["TypesOfSource"] = (xlob.xl_type_of_source_id[b])
        # print self.data
        ob.json_data()
        self.api_types_of_source = self.count
        # print "api_types_of_source Count is %s " %self.api_types_of_source


        # Candidate Source
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"SourceId":"8396"},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["SourceId"] = int((xlob.xl_source_id[b]))
        # print self.data
        ob.json_data()
        self.api_candidate_source = self.count
        # print "api_candidate_source Count is %s " %self.api_candidate_source

        # Candidate College
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"CollegeIds":["1637"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["CollegeIds"] = [int((xlob.xl_college_id[b]))]
        # print self.data
        ob.json_data()
        self.api_candidate_college = self.count
        # print "api_candidate_college Count is %s " %self.api_candidate_college

        # Candidate Degree
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"DegreeIds":["1533"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["DegreeIds"] = [int((xlob.xl_degree_id[b]))]
        # print self.data
        ob.json_data()
        self.api_candidate_degree = self.count
        # print "api_candidate_degree Count is %s " %self.api_candidate_degree

        # Candidate Department
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"BranchIds":["255"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["BranchIds"] = [int((xlob.xl_department_id[b]))]
        # print self.data
        ob.json_data()
        self.api_candidate_department = self.count
        # print "api_candidate_department Count is %s " %self.api_candidate_department

        # Candidate YOP
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"YearEnd":["12790"]},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["YearEnd"] = [int((xlob.xl_yop_id[b]))]
        # print self.data
        ob.json_data()
        self.api_candidate_yop = self.count
        # print "api_candidate_yop Count is %s " %self.api_candidate_yop

        # Candidate Percentage
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"PercentageFrom":50,"PercentageTo":100},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["PercentageFrom"] = int((xlob.xl_percentage_from[b]))
        self.data["CandidateFilters"]["PercentageTo"] = int((xlob.xl_percentage_to[b]))
        # print self.data
        ob.json_data()
        self.api_candidate_percentage = self.count
        # print "api_candidate_percentage Count is %s " %self.api_candidate_percentage

        # Candidate Marital Status
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"MaritalStatus":"1"},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["MaritalStatus"] = int((xlob.xl_marital_status_id[b]))

        ob.json_data()
        self.api_candidate_marital_status = self.count
        # print "api_candidate_marital_status Count is %s " %self.api_candidate_marital_status

        # Candidate Sourcer
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"Sourcer":"37119"},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["Sourcer"] = int((xlob.xl_sourcer_id[b]))
        ob.json_data()
        self.api_candidate_sourcer = self.count
        # print "api_candidate_sourcer Count is %s " %self.api_candidate_sourcer

        # Candidate Created on
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"CreatedFrom":"2017-09-30T18:30:00.000Z","CreatedTo":"2018-10-09T18:30:00.000Z"},
                    "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["CreatedFrom"] = xlob.xl_created_on_from[b]
        self.data["CandidateFilters"]["CreatedTo"] = xlob.xl_created_on_to[b]
        # print self.data
        ob.json_data()
        self.api_candidate_created_on = self.count
        # print "api_candidate_created_on Count is %s " %self.api_candidate_created_on

        # Candidate Modified on
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"ModifiedFrom":"2017-09-30T18:30:00.000Z","ModifiedTo":"2018-10-09T18:30:00.000Z"},
                     "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["ModifiedFrom"] = xlob.xl_modified_on_from[b]
        self.data["CandidateFilters"]["ModifiedTo"] = xlob.xl_modified_on_to[b]
        # print self.data
        ob.json_data()
        self.api_candidate_modified_on = self.count
        # print "api_candidate_modified_on Count is %s " %self.api_candidate_modified_on

        # Candidate Created by
        self.data ={"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                    "CandidateFilters":{"CreatedBy":"37119"},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["CreatedBy"] = int((xlob.xl_created_by_id[b]))
        ob.json_data()
        self.api_candidate_created_by = self.count
        # print "api_candidate_created_by Count is %s " %self.api_candidate_created_by

        # Candidate is utilized
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"CandidateUtilization":False},
                    "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["CandidateUtilization"] = (xlob.xl_is_utilized[b])
        ob.json_data()
        self.api_candidate_is_utilized = self.count
        # print "api_candidate_is_utilized Count is %s " %self.api_candidate_is_utilized


        # Candidate Gender
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},"CandidateFilters":{"Gender":"1"},
                     "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["Gender"] = int((xlob.xl_is_gender_id[b]))
        ob.json_data()
        self.api_candidate_gender = self.count
        # print "api_candidate_gender Count is %s " %self.api_candidate_gender

        # Candidate USN
        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},"CandidateFilters":{"USN":"RPOTEST0002"},
                     "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["USN"] = (xlob.xl_usn[b])
        ob.json_data()
        self.api_candidate_usn = self.count
        # print "api_candidate_usn Count is %s " %self.api_candidate_usn

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"Text1":"T1"},"CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["Text1"] = (xlob.xl_text1[b])
        ob.json_data()
        self.api_text1 = self.count
        # print "API Text1 Count is %s " %self.api_text1

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Text2": "t2"}, "CandidateCustomFilters": {}, "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Text2"] = (xlob.xl_text2[b])
        ob.json_data()
        self.api_text2 = self.count
        # print "API Text2 Count is %s " %self.api_text2

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Text3": "t3"}, "CandidateCustomFilters": {}, "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Text3"] = (xlob.xl_text3[b])
        ob.json_data()
        self.api_text3 = self.count
        # print "API Text3 Count is %s " %self.api_text3

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Text4": "t4"}, "CandidateCustomFilters": {}, "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Text4"] = (xlob.xl_text4[b])
        ob.json_data()
        self.api_text4 = self.count
        # print "API Text4 Count is %s " %self.api_text4

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Text5": "t5"}, "CandidateCustomFilters": {}, "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Text5"] = (xlob.xl_text5[b])
        ob.json_data()
        self.api_text5 = self.count
        # print "API Text5 Count is %s " %self.api_text5

        self.data ={"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                    "CandidateFilters":{},"CandidateCustomFilters":{"Text6":"t6"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text6"] = (xlob.xl_text6[b])
        ob.json_data()
        self.api_text6 = self.count
        # print "API Text6 Count is %s " % self.api_text6

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text7":"t7"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text7"] = (xlob.xl_text7[b])
        ob.json_data()
        self.api_text7 = self.count
        # print "API Text7 Count is %s " % self.api_text7

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text8":"t8"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text8"] = (xlob.xl_text8[b])
        ob.json_data()
        self.api_text8 = self.count
        # print "API Text8 Count is %s " % self.api_text8

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text9":"t9"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text9"] = (xlob.xl_text9[b])
        ob.json_data()
        self.api_text9 = self.count
        # print "API Text9 Count is %s " % self.api_text9

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text10":"t10"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text10"] = (xlob.xl_text10[b])
        self.api_text10 = self.count
        # print "API Text10 Count is %s " % self.api_text10

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text11":"t11"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text11"] = (xlob.xl_text11[b])
        ob.json_data()
        self.api_text11 = self.count
        # print "API Text11 Count is %s " % self.api_text11

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text12":"t12"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text12"] = (xlob.xl_text12[b])
        ob.json_data()
        self.api_text12 = self.count
        # print "API Text12 Count is %s " % self.api_text12

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text13":"t13"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text13"] = (xlob.xl_text13[b])
        ob.json_data()
        self.api_text13 = self.count
        # print "API Text13 Count is %s " % self.api_text13

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text14":"t14"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text14"] = (xlob.xl_text14[b])
        ob.json_data()
        self.api_text14 = self.count
        # print "API Text14 Count is %s " % self.api_text14

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Text15":"t15"},"IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Text15"] = (xlob.xl_text15[b])
        ob.json_data()
        self.api_text15 = self.count
        # print "API Text15 Count is %s " % self.api_text15

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"Integer1":["50"]},"CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["Integer1"] = [int((xlob.xl_integer1[b]))]
        ob.json_data()
        self.api_integer1 = self.count
        # print "API integer1 Count is %s " % self.api_integer1

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Integer2": ["34"]}, "CandidateCustomFilters": {},"IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Integer2"] = [int((xlob.xl_integer2[b]))]
        ob.json_data()
        self.api_integer2 = self.count
        # print "API integer2 Count is %s " % self.api_integer2

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Integer3": ["1533"]}, "CandidateCustomFilters": {},
                     "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Integer3"] = [int((xlob.xl_integer3[b]))]
        ob.json_data()
        self.api_integer3 = self.count
        # print "API integer3 Count is %s " % self.api_integer3

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Integer4": ["108"]}, "CandidateCustomFilters": {},
                     "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Integer4"] = [int((xlob.xl_integer4[b]))]
        ob.json_data()
        self.api_integer4 = self.count
        # print "API integer4 Count is %s " % self.api_integer4

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {"Integer5": ["325"]}, "CandidateCustomFilters": {},
                     "IsNotCacheRequired": False}
        self.data["CandidateFilters"]["Integer5"] = [int((xlob.xl_integer5[b]))]
        ob.json_data()
        self.api_integer5 = self.count
        # print "API integer5 Count is %s " % self.api_integer5

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Integer6":["255"]},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Integer6"] = [int((xlob.xl_integer6[b]))]
        ob.json_data()
        self.api_integer6 = self.count
        # print "API integer1 Count is %s " % self.api_integer6

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Integer7":["1809"]},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Integer7"] = [int((xlob.xl_integer7[b]))]
        ob.json_data()
        self.api_integer7 = self.count
        # print "API integer7 Count is %s " % self.api_integer7

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Integer8":["2152"]},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Integer8"] = [int((xlob.xl_integer8[b]))]
        ob.json_data()
        self.api_integer8 = self.count
        # print "API integer8 Count is %s " % self.api_integer8

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Integer9":["244"]},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Integer9"] = [int((xlob.xl_integer9[b]))]
        ob.json_data()
        self.api_integer9 = self.count
        # print "API integer9 Count is %s " % self.api_integer9

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{},"CandidateCustomFilters":{"Integer10":["3054"]},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["Integer10"] = [int((xlob.xl_integer10[b]))]
        ob.json_data()
        self.api_integer10 = self.count
        # print "API integer10 Count is %s " % self.api_integer10

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {}, "CandidateCustomFilters": {"Integer11": ["1968"]},
                     "IsNotCacheRequired": False}
        self.data["CandidateCustomFilters"]["Integer11"] = [int((xlob.xl_integer11[b]))]
        ob.json_data()
        self.api_integer11 = self.count
        # print "API integer11 Count is %s " % self.api_integer11

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {}, "CandidateCustomFilters": {"Integer12": ["2995"]},
                     "IsNotCacheRequired": False}
        self.data["CandidateCustomFilters"]["Integer12"] = [int((xlob.xl_integer12[b]))]
        ob.json_data()
        self.api_integer12 = self.count
        # print "API integer12 Count is %s " % self.api_integer12

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {}, "CandidateCustomFilters": {"Integer13": ["1637"]},
                     "IsNotCacheRequired": False}
        self.data["CandidateCustomFilters"]["Integer13"] = [int((xlob.xl_integer13[b]))]
        ob.json_data()
        self.api_integer13 = self.count
        # print "API integer13 Count is %s " % self.api_integer13

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {}, "CandidateCustomFilters": {"Integer14": ["38"]},
                     "IsNotCacheRequired": False}
        self.data["CandidateCustomFilters"]["Integer14"] = [int((xlob.xl_integer14[b]))]
        ob.json_data()
        self.api_integer14 = self.count
        # print "API integer14 Count is %s " % self.api_integer14

        self.data = {"PagingCriteria": {"ObjectState": 0, "IsCountOnly": True},
                     "CandidateFilters": {}, "CandidateCustomFilters": {"Integer15": ["16625"]},
                     "IsNotCacheRequired": False}
        self.data["CandidateCustomFilters"]["Integer15"] = [int((xlob.xl_integer15[b]))]
        ob.json_data()
        self.api_integer15 = self.count
        # print "API integer15 Count is %s " % self.api_integer15

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"TextArea1":"ta1"},"CandidateCustomFilters":{},
                     "IsNotCacheRequired":False}
        self.data["CandidateFilters"]["TextArea1"] = (xlob.xl_textarea1[b])
        ob.json_data()
        self.api_textarea1 = self.count
        # print "API textarea1 Count is %s " % self.api_textarea1

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"TextArea2":"ta2"},"CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["TextArea2"] = (xlob.xl_textarea2[b])
        ob.json_data()
        self.api_textarea2 = self.count
        # print "API textarea2 Count is %s " % self.api_textarea2

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"TextArea3":"ta3"},"CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["TextArea3"] = (xlob.xl_textarea3[b])
        ob.json_data()
        self.api_textarea3 = self.count
        # print "API textarea3 Count is %s " % self.api_textarea3

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"TextArea4":"ta4"},"CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["TextArea4"] = (xlob.xl_textarea4[b])
        ob.json_data()
        self.api_textarea4= self.count
        # print "API textarea4 Count is %s " % self.api_textarea4

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"DateCustomField1":"2017-10-30T18:30:00.000Z"},
                     "CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["DateCustomField1"] = (xlob.xl_date1[b])
        ob.json_data()
        self.api_date1= self.count
        # print "API date1 Count is %s " % self.api_date1

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},
                     "CandidateFilters":{"DateCustomField2":"2017-10-31T18:30:00.000Z"},
                     "CandidateCustomFilters":{},"IsNotCacheRequired":False}
        self.data["CandidateFilters"]["DateCustomField2"] = (xlob.xl_date2[b])
        ob.json_data()
        self.api_date2 = self.count
        # print "API date2 Count is %s " % self.api_date2

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},"CandidateFilters":{},
                     "CandidateCustomFilters":{"DateCustomField3":"2017-10-11T18:30:00.000Z"},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["DateCustomField3"] = (xlob.xl_date3[b])

        ob.json_data()
        self.api_date3 = self.count
        # print "API date3 Count is %s " % self.api_date3

        self.data = {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},"CandidateFilters":{},
                     "CandidateCustomFilters":{"DateCustomField4":"2017-11-12T18:30:00.000Z"},
                     "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["DateCustomField4"] = (xlob.xl_date4[b])
        ob.json_data()
        self.api_date4 = self.count
        # print "API date4 Count is %s " % self.api_date4

        self.data =  {"PagingCriteria":{"ObjectState":0,"IsCountOnly":True},"CandidateFilters":{},
                      "CandidateCustomFilters":{"DateCustomField5":"2017-10-13T18:30:00.000Z"},
                      "IsNotCacheRequired":False}
        self.data["CandidateCustomFilters"]["DateCustomField5"] = (xlob.xl_date5[b])
        ob.json_data()
        self.api_date5 = self.count
        # print "API date5 Count is %s " % self.api_date5

    def Compare_Values(self, search_data,api_value, db_value,mess):
        self.ws.write(self.rowsize, self.a, search_data, self.__style3)
        if (api_value == db_value):
            self.ws.write(self.rowsize+1, self.a, api_value, self.__style1)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style1)
            self.a = self.a + 1
            # print self.a
        elif (api_value == '400000000000000'):
            self.ws.write(self.rowsize+1, self.a, 'API is Throwing Error', self.__style2)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style2)
            self.a = self.a + 1
            # print self.a
        else:
            self.ws.write(self.rowsize+1, self.a, api_value, self.__style2)
            self.ws.write(self.rowsize+2, self.a, db_value, self.__style2)
            self.a = self.a + 1
            print self.mess
            print "api Value is %s"% api_value
            # print type(api_value)
            print "DB Value is %s" % db_value
            # print type(api_value)

    def validation_MS_and_UI1(self):
        now = datetime.datetime.now()
        self.rowsize = 1
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        # print self.__current_DateTime
        # print type(self.__current_DateTime)
        self.__style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on')
        self.__style1 = xlwt.easyxf('font: name Times New Roman, color-index black, bold off')
        self.__style2 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
        self.__style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on')
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Candidate_Search')
        self.ws.write(0, 0, 'Type', self.__style0)


        self.ws.write(0, 1, 'Candidate Name', self.__style0)
        self.ws.write(0, 2, 'Email', self.__style0)
        self.ws.write(0, 3, 'Mobile', self.__style0)
        self.ws.write(0, 4, 'Department', self.__style0)
        self.ws.write(0, 5, 'Job Role', self.__style0)
        self.ws.write(0, 6, 'Status', self.__style0)
        self.ws.write(0, 7, 'Current Location', self.__style0)
        self.ws.write(0, 8, 'Experience', self.__style0)
        self.ws.write(0, 9, 'Current Employer', self.__style0)
        self.ws.write(0, 10, 'Expertise', self.__style0)
        self.ws.write(0, 11, 'Type of source', self.__style0)
        self.ws.write(0, 12, 'Source', self.__style0)

        self.ws.write(0, 13, 'College', self.__style0)
        self.ws.write(0, 14, 'Degree', self.__style0)
        self.ws.write(0, 15, 'Department', self.__style0)
        self.ws.write(0, 16, 'YOP', self.__style0)
        self.ws.write(0, 17, 'Percentage', self.__style0)

        self.ws.write(0, 18, 'Marital Status', self.__style0)
        self.ws.write(0, 19, 'Created On', self.__style0)
        self.ws.write(0, 20, 'Modified On', self.__style0)
        self.ws.write(0, 21, 'Sourcer', self.__style0)
        self.ws.write(0, 22, 'Created By', self.__style0)
        self.ws.write(0, 23, 'Is_Utilized', self.__style0)

        self.ws.write(0, 24, 'Gender', self.__style0)
        self.ws.write(0, 25, 'USN', self.__style0)


        self.ws.write(0, 26, 'Text1', self.__style0)
        self.ws.write(0, 27, 'Text2', self.__style0)
        self.ws.write(0, 28, 'Text3', self.__style0)
        self.ws.write(0, 29, 'Text4', self.__style0)
        self.ws.write(0, 30, 'Text5', self.__style0)
        self.ws.write(0, 31, 'Text6', self.__style0)
        self.ws.write(0, 32, 'Text7', self.__style0)
        self.ws.write(0, 33, 'Text8', self.__style0)
        self.ws.write(0, 34, 'Text9', self.__style0)
        self.ws.write(0, 35, 'Text10', self.__style0)
        self.ws.write(0, 36, 'Text11', self.__style0)
        self.ws.write(0, 37, 'Text12', self.__style0)
        self.ws.write(0, 38, 'Text13', self.__style0)
        self.ws.write(0, 39, 'Text14', self.__style0)
        self.ws.write(0, 40, 'Text15', self.__style0)

        self.ws.write(0, 41, 'Integer1', self.__style0)
        self.ws.write(0, 42, 'Integer2', self.__style0)
        self.ws.write(0, 43, 'Integer3', self.__style0)
        self.ws.write(0, 44, 'Integer4', self.__style0)
        self.ws.write(0, 45, 'Integer5', self.__style0)
        self.ws.write(0, 46, 'Integer6', self.__style0)
        self.ws.write(0, 47, 'Integer7', self.__style0)
        self.ws.write(0, 48, 'Integer8', self.__style0)
        self.ws.write(0, 49, 'Integer9', self.__style0)
        self.ws.write(0, 50, 'Integer10', self.__style0)
        self.ws.write(0, 51, 'Integer11', self.__style0)
        self.ws.write(0, 52, 'Integer12', self.__style0)
        self.ws.write(0, 53, 'Integer13', self.__style0)
        self.ws.write(0, 54, 'Integer14', self.__style0)
        self.ws.write(0, 55, 'Integer15', self.__style0)

        self.ws.write(0, 56, 'TextArea1', self.__style0)
        self.ws.write(0, 57, 'TextArea2', self.__style0)
        self.ws.write(0, 58, 'TextArea3', self.__style0)
        self.ws.write(0, 59, 'TextArea4', self.__style0)

        self.ws.write(0, 60, 'Date1', self.__style0)
        self.ws.write(0, 61, 'Date2', self.__style0)
        self.ws.write(0, 62, 'Date3', self.__style0)
        self.ws.write(0, 63, 'Date4', self.__style0)
        self.ws.write(0, 64, 'Date5', self.__style0)

        # self.ws.write(0, 65, 'Truefalse1', self.__style0)
        # self.ws.write(0, 66, 'Truefalse2', self.__style0)
        # self.ws.write(0, 67, 'Truefalse3', self.__style0)
        # self.ws.write(0, 68, 'Truefalse4', self.__style0)
        # self.ws.write(0, 69, 'Truefalse5', self.__style0)


    def validation_MS_and_UI2(self, b):
        self.a = 1
        # print self.rowsize
        # print self.a
        self.ws.write(self.rowsize, 0, 'Search Keyword', self.__style0)
        self.ws.write(self.rowsize + 1, 0, 'API Count', self.__style0)
        self.ws.write(self.rowsize + 2, 0, 'DB Count', self.__style0)
        # print "Compare method is started"
        self.mess = 'Candidate Name is  not matched'
        self.Compare_Values(xlob.xl_candidate_name[b],ob.api_candidate_name, C4.db_candidate_name,self.mess)
        self.mess = 'Email is  not matched'
        self.Compare_Values(xlob.xl_email[b],ob.api_candidate_email, C4.db_candidate_email, self.mess)
        self.mess = 'Mobile is not matched'
        self.Compare_Values(xlob.xl_mobile[b],ob.api_candidate_mobile, C4.db_candidate_mobile, self.mess)

        self.search_data = C4.db_job_role_department_name
        self.mess = 'Job Department is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_job_department,C4.db_candidate_job_department , self.mess)

        self.search_data = C4.db_job_name
        self.mess = 'Jobrole applicants is not matched'
        self.Compare_Values(self.search_data,ob.api_jobrole_applicants, C4.db_candidate_jobapplicant, self.mess)

        self.search_data = C4.db_stage_status_name
        self.mess = 'Stage Status is not matched'
        self.Compare_Values(self.search_data ,ob.api_candidate_stage_status, C4.db_candidate_stage_status, self.mess)


        self.search_data = C4.db_location_name
        self.mess = 'Current Location is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_current_location, C4.db_candidate_current_location, self.mess)

        self.mess = 'Experience is not matched'
        experience = "%s to %s Year" %(str(xlob.xl_experience_from[b]),str(xlob.xl_experience_to[b]))
        self.Compare_Values(experience,ob.api_candidate_experience, C4.db_candidate_total_experience, self.mess)

        self.search_data = C4.db_employer_name
        self.mess = 'Current Organization / Employer is not matched'
        self.Compare_Values(self.search_data ,ob.api_candidate_organization, C4.db_candidate_current_employer, self.mess)

        self.search_data = C4.db_expertise_name
        self.mess = 'Expertise is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_expertise, C4.db_candidate_expertise, self.mess)
        self.mess = 'Type of source is not matched'
        self.Compare_Values(xlob.xl_type_of_source_id[b],ob.api_types_of_source, C4.db_types_of_source, self.mess)

        self.search_data = C4.db_source_name
        self.mess = 'source is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_source, C4.db_source, self.mess)

        self.search_data = C4.db_college_name
        self.mess = 'College is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_college, C4.db_college, self.mess)

        self.search_data = C4.db_degree_name
        self.mess = 'Degree is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_degree, C4.db_degree, self.mess)

        self.search_data = C4.db_department_name
        self.mess = 'Department is not matched'
        self.Compare_Values(self.search_data ,ob.api_candidate_department, C4.db_department, self.mess)

        self.search_data =C4.db_yop_name
        self.mess = 'YOP is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_yop, C4.db_yop, self.mess)

        self.mess = 'Percentage is not matched'
        percentage = "%s to %s" % (str(xlob.xl_percentage_from[b]), str(xlob.xl_percentage_to[b]))
        self.Compare_Values(percentage,ob.api_candidate_percentage, C4.db_percentage, self.mess)

        self.mess = 'Marital Status is not matched'
        self.Compare_Values(xlob.xl_marital_status_id[b],ob.api_candidate_marital_status, C4.db_marital, self.mess)

        self.mess = 'Created on is not matched'
        created_on = "%s to %s " %((xlob.xl_created_on_from[b]),(xlob.xl_created_on_to[b]))
        self.Compare_Values(created_on,ob.api_candidate_created_on, C4.db_created_on, self.mess)

        self.mess = 'Modefied on is not matched'
        modified_on = "%s to %s " % ((xlob.xl_modified_on_from[b]), (xlob.xl_modified_on_to[b]))
        self.Compare_Values(modified_on,ob.api_candidate_modified_on, C4.db_modefied_on, self.mess)

        self.search_data = C4.db_sourcer_name
        self.mess = 'Sourcer is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_sourcer, C4.db_sourcer, self.mess)

        self.search_data = C4.db_createdby_name
        self.mess = 'Created by  is not matched'
        self.Compare_Values(self.search_data,ob.api_candidate_created_by, C4.db_created_by, self.mess)

        self.mess = 'Is Utilized is not matched'
        self.Compare_Values(xlob.xl_is_utilized[b],ob.api_candidate_is_utilized,C4.db_is_utilized, self.mess)
        self.mess = 'Gender is not matched'
        self.Compare_Values(xlob.xl_is_gender_id[b],ob.api_candidate_gender, C4.db_is_gender, self.mess)
        self.mess = 'USN is not matched'
        self.Compare_Values(xlob.xl_usn[b],ob.api_candidate_usn, C4.db_is_usn, self.mess)

        self.mess = 'Text1 is not matched'
        self.Compare_Values(xlob.xl_text1[b],ob.api_text1, C4.db_text1, self.mess)
        self.mess = 'Text2 is not matched'
        self.Compare_Values(xlob.xl_text2[b],ob.api_text2, C4.db_text2, self.mess)
        self.mess = 'Text3 is not matched'
        self.Compare_Values(xlob.xl_text3[b],ob.api_text3, C4.db_text3, self.mess)
        self.mess = 'Text4 is not matched'
        self.Compare_Values(xlob.xl_text4[b],ob.api_text4, C4.db_text4, self.mess)
        self.mess = 'Text5 is not matched'
        self.Compare_Values(xlob.xl_text5[b],ob.api_text5, C4.db_text5, self.mess)
        self.mess = 'Text6 is not matched'
        self.Compare_Values(xlob.xl_text6[b],ob.api_text6, C4.db_text6, self.mess)
        self.mess = 'Text7 is not matched'
        self.Compare_Values(xlob.xl_text7[b],ob.api_text7, C4.db_text7, self.mess)
        self.mess = 'Text8 is not matched'
        self.Compare_Values(xlob.xl_text8[b],ob.api_text8, C4.db_text8, self.mess)
        self.mess = 'Text9 is not matched'
        self.Compare_Values(xlob.xl_text9[b],ob.api_text9, C4.db_text9, self.mess)
        self.mess = 'Text10 is not matched'
        self.Compare_Values(xlob.xl_text10[b],ob.api_text10, C4.db_text10, self.mess)
        self.mess = 'Text11 is not matched'
        self.Compare_Values(xlob.xl_text11[b],ob.api_text11, C4.db_text11, self.mess)
        self.mess = 'Text12 is not matched'
        self.Compare_Values(xlob.xl_text12[b],ob.api_text12, C4.db_text12, self.mess)
        self.mess = 'Text13 is not matched'
        self.Compare_Values(xlob.xl_text13[b],ob.api_text13, C4.db_text13, self.mess)
        self.mess = 'Text14 is not matched'
        self.Compare_Values(xlob.xl_text14[b],ob.api_text14, C4.db_text14, self.mess)
        self.mess = 'Text15 is not matched'
        self.Compare_Values(xlob.xl_text15[b],ob.api_text15, C4.db_text15, self.mess)

        self.search_data = C4.db_integer1_name
        self.mess = 'Integer1 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer1, C4.db_integer1, self.mess)
        self.search_data = C4.db_integer2_name
        self.mess = 'Integer2 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer2, C4.db_integer2, self.mess)
        self.search_data = C4.db_integer3_name
        self.mess = 'Integer3 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer3, C4.db_integer3, self.mess)
        self.search_data = C4.db_integer4_name
        self.mess = 'Integer4 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer4, C4.db_integer4, self.mess)
        self.search_data = C4.db_integer5_name
        self.mess = 'Integer5 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer5, C4.db_integer5, self.mess)
        self.search_data = C4.db_integer6_name
        self.mess = 'Integer6 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer6, C4.db_integer6, self.mess)
        self.search_data = C4.db_integer7_name
        self.mess = 'Integer7 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer7, C4.db_integer7, self.mess)
        self.search_data = C4.db_integer8_name
        self.mess = 'Integer8 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer8, C4.db_integer8, self.mess)
        self.search_data = C4.db_integer9_name
        self.mess = 'Integer9 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer9, C4.db_integer9, self.mess)
        self.search_data = C4.db_integer10_name
        self.mess = 'Integer10 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer10, C4.db_integer10, self.mess)
        self.search_data = C4.db_integer11_name
        self.mess = 'Integer11 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer11, C4.db_integer11, self.mess)
        self.search_data = C4.db_integer12_name
        self.mess = 'Integer12 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer12, C4.db_integer12, self.mess)
        self.search_data = C4.db_integer13_name
        self.mess = 'Integer13 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer13, C4.db_integer13, self.mess)
        self.search_data = C4.db_integer14_name
        self.mess = 'Integer14 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer14, C4.db_integer14, self.mess)
        self.search_data = C4.db_integer15_name
        self.mess = 'Integer15 is not matched'
        self.Compare_Values(self.search_data,ob.api_integer15, C4.db_integer15, self.mess)

        self.mess = 'Textarea1 is not matched'
        self.Compare_Values(xlob.xl_textarea1[b],ob.api_textarea1, C4.db_textarea1, self.mess)
        self.mess = 'Textarea2 is not matched'
        self.Compare_Values(xlob.xl_textarea2[b],ob.api_textarea2, C4.db_textarea2, self.mess)
        self.mess = 'Textarea3 is not matched'
        self.Compare_Values(xlob.xl_textarea3[b],ob.api_textarea3, C4.db_textarea3, self.mess)
        self.mess = 'Textarea4 is not matched'
        self.Compare_Values(xlob.xl_textarea4[b],ob.api_textarea4, C4.db_textarea4, self.mess)

        self.mess = 'Date1 is not matched'
        self.Compare_Values(xlob.xl_date1[b],ob.api_date1, C4.db_date1, self.mess)
        self.mess = 'Date2 is not matched'
        self.Compare_Values(xlob.xl_date2[b],ob.api_date2, C4.db_date2, self.mess)
        self.mess = 'Date3 is not matched'
        self.Compare_Values(xlob.xl_date3[b],ob.api_date3, C4.db_date3, self.mess)
        self.mess = 'Date4 is not matched'
        self.Compare_Values(xlob.xl_date4[b],ob.api_date4, C4.db_date4, self.mess)
        self.mess = 'Date5 is not matched'
        self.Compare_Values(xlob.xl_date5[b],ob.api_date5, C4.db_date5, self.mess)
        self.wb_Result.save(
                    '/home/muttumurgan/Desktop/PythonWorkingScripts/OutputData/CRPO/Search/API_DB_Search_Verification(' + self.__current_DateTime + ').xls')
        self.rowsize = self.rowsize +4

C4 = AMS_DB_Data()
ob = api_data()
print len(xlob.xl_candidate_id)
tot_count = len(xlob.xl_candidate_id)
ob.validation_MS_and_UI1()

for b in range(0,tot_count):
    print b
    print "Candidate Search Script Started....."
    C4.ams_Query(b)
    ob.api_main(b)
    ob.validation_MS_and_UI2(b)