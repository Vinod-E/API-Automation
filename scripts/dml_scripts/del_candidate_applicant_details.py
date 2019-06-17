import xlrd
from hpro_automation import (output_paths, db_login)


class DeleteCandidateApplicant(db_login.DBConnection):
    def __init__(self):
        super(DeleteCandidateApplicant, self).__init__()
        self.db_connection()

        # ------------------------
        # Initialising Excel Data
        # ------------------------
        self.xl_candidateId = []

    def candidate_excel_data(self):

        workbook = xlrd.open_workbook(output_paths.DMLOutput['candidate_output_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(2, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[2]:
                self.xl_candidateId.append(int(rows[2]))

        xl_candidateid = tuple(self.xl_candidateId)

        candidate_query1 = "DELETE FROM duplicate_candidates_infos where candidate_id in{} and tenant_id=1787;" \
            .format(xl_candidateid)
        self.cursor.execute(candidate_query1)
        print(candidate_query1)

        candidate_query2 = "DELETE FROM test_users WHERE candidate_id in{};".format(xl_candidateid)
        self.cursor.execute(candidate_query2)
        print(candidate_query2)

        candidate_query3 = "Delete from applicant_status_items where applicantstatus_id in" \
                           " (select id from applicant_statuss where recruitevent_id=4697 and " \
                           "tenant_id=1787 and is_deleted=0)and tenant_id=1787;"
        self.cursor.execute(candidate_query3)
        print(candidate_query3)

        candidate_query4 = "DELETE FROM applicant_statuss " \
                           "WHERE recruitevent_id=4697 and tenant_id=1787 and is_deleted=0;"
        self.cursor.execute(candidate_query4)
        print(candidate_query4)

        candidate_query5 = "DELETE from test_user_applicants where testuser_id not in (select id from test_users);"
        self.cursor.execute(candidate_query5)
        print(candidate_query5)

        candidate_query6 = "DELETE FROM appserver_core.candidates " \
                           "WHERE tenant_id=1787 and id in{};".format(xl_candidateid)
        self.cursor.execute(candidate_query6)
        print(candidate_query6)


Object = DeleteCandidateApplicant()
Object.candidate_excel_data()
Object.connection.commit()
Object.connection.close()
