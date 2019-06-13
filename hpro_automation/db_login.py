from hpro_automation.identity import credentials
import pymysql


class DBConnection(object):

    def __init__(self):
        super(DBConnection, self).__init__()
        self.user = str(input("DB name:: "))

    def db_connection(self):

        self.DB_data = credentials.db_details[self.user]
        self.connection = pymysql.connect(**self.DB_data)
        self.cursor = self.connection.cursor()
