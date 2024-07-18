from hpro_automation import (input_paths, db_login)
import xlrd


class DeleteActivities(db_login.DBConnection):

    def __init__(self):
        super(DeleteActivities, self).__init__()
        self.db_connection()

        # -----------------------
        # Initialising the values
        # -----------------------
        self.xl_candidateId = []
        self.candidatestaffingprofile_id = []
        self.candidate_user_id = []
        self.assigned_task_id = []
        self.filled_form_id = []

    def activity_excel_data(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['Activity_C_back_Input_sheet'])
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[0]:
                self.xl_candidateId.append(int(rows[0]))

    def fetch_candidate_table(self, loop):
        query = "select candidatestaffingprofile_id, candidate_user_id from candidates " \
                "where id = {};" .format(self.xl_candidateId[loop])
        print(query)
        self.cursor.execute(query)
        records = self.cursor.fetchall()
        print("Total number of rows ::", self.cursor.rowcount)
        for row in records:
            if row[0] is not None:
                self.candidatestaffingprofile_id.append(int(row[0]))
            else:
                self.candidatestaffingprofile_id.append(0)

            if row[1] is not None:
                self.candidate_user_id.append(int(row[1]))
            else:
                self.candidate_user_id.append(0)

    def fetch_assign_task_table(self, loop):
        query = "select id, filled_form_id from assigned_tasks " \
                "where candidate_id = {};" .format(self.xl_candidateId[loop])
        print(query)
        self.cursor.execute(query)
        records = self.cursor.fetchall()
        print("Total number of rows ::", self.cursor.rowcount)
        for row in records:
            if row[0] is not None:
                self.assigned_task_id.append(int(row[0]))
            else:
                self.assigned_task_id.append(0)

            if row[1] is not None:
                self.filled_form_id.append(int(row[1]))
            else:
                self.filled_form_id.append(0)

    def making_rows_delete_null(self):

        xl_candidates = tuple(self.xl_candidateId)
        db_users = tuple(self.candidate_user_id)
        db_staffing_profile = tuple(self.candidatestaffingprofile_id)
        db_assigned_task_id = tuple(self.assigned_task_id)
        db_filled_form_id = tuple(self.filled_form_id)

        # ------------------------------------ Candidate table ---------------------------------------------------------
        query = "UPDATE candidates SET " \
                "current_activity = NULL, " \
                "activity_status = NULL, " \
                "candidate_user_id = NULL, " \
                "candidatestaffingprofile_id = NULL " \
                "WHERE id in {};".format(xl_candidates)
        self.cursor.execute(query)
        print(query)

        # -------------------------------------- Users table -----------------------------------------------------------
        query1 = "DELETE FROM users WHERE  id in {};".format(db_users)
        self.cursor.execute(query1)
        print(query1)

        # -------------------------------- candidate_staffing_profiles -------------------------------------------------
        query2 = "DELETE FROM candidate_staffing_profiles WHERE id in{};".format(db_staffing_profile)
        self.cursor.execute(query2)
        print(query2)

        # ------------------------------------ candidate_activitys -----------------------------------------------------
        query3 = "DELETE FROM candidate_activitys " \
                 "WHERE candidate_id in ({});".format(', '.join(map(str, xl_candidates)))
        self.cursor.execute(query3)
        print(query3)

        # ------------------------------------ assigned_task_historys --------------------------------------------------
        query4 = "DELETE FROM assigned_task_historys" \
                 " WHERE assignedtask_id in ({});".format(','.join(map(str, db_assigned_task_id)))
        self.cursor.execute(query4)
        print(query4)

        # ---------------------------------------- assigned_tasks ------------------------------------------------------
        query5 = "DELETE FROM assigned_tasks WHERE candidate_id in {};".format(xl_candidates)
        self.cursor.execute(query5)
        print(query5)

        # ------------------------------------ assigned_task_historys --------------------------------------------------
        query6 = "DELETE FROM filled_form_contents " \
                 "WHERE filledform_id in ({});".format(','.join(map(str, db_filled_form_id)))
        self.cursor.execute(query6)
        print(query6)

        # ------------------------------------ assigned_task_historys --------------------------------------------------
        query7 = "DELETE FROM filled_forms WHERE id in ({});".format(','.join(map(str, db_filled_form_id)))
        self.cursor.execute(query7)
        print(query7)


Object = DeleteActivities()
Object.activity_excel_data()
Total_count = len(Object.xl_candidateId)
print("Number Of Rows ::", Total_count)
for looping in range(0, Total_count):
    print("Iteration Count is ::", looping)
    Object.fetch_candidate_table(looping)
    Object.fetch_assign_task_table(looping)
Object.making_rows_delete_null()

print(Object.candidatestaffingprofile_id)
print(Object.candidate_user_id)
print(Object.assigned_task_id)
print(Object.filled_form_id)
Object.connection.commit()
Object.connection.close()
