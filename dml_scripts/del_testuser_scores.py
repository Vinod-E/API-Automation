import xlrd
from hpro_automation import (input_paths, db_login)


class TestUserScores(db_login.DBConnection):
    def __init__(self):
        super(TestUserScores, self).__init__()
        self.db_connection('amsin')

        # ------------------------
        # Initialising Excel Data
        # ------------------------
        self.xl_testuserId = []

    def test_user_excel_data(self):
        workbook = xlrd.open_workbook(input_paths.inputpaths['Uploadscore_Input_sheet'])
        sheet1 = workbook.sheet_by_index(1)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)
            if rows[0]:
                self.xl_testuserId.append(int(rows[0]))

        xl_testuserid1 = tuple(self.xl_testuserId)
        testuser_query = "delete from candidate_scores where testuser_id in{};" \
            .format(xl_testuserid1)
        self.cursor.execute(testuser_query)
        print(testuser_query)

        testuser_query1 = "UPDATE test_users SET total_score=Null, percentage=Null, status=Null WHERE id in {};" \
            .format(xl_testuserid1)
        self.cursor.execute(testuser_query1)
        print(testuser_query1)


Object = TestUserScores()
Object.test_user_excel_data()
Object.connection.commit()
Object.connection.close()
