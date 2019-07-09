from hpro_automation.identity import credentials
import pymysql


class DBConnection(object):

    def __init__(self):
        super(DBConnection, self).__init__()

    def db_connection(self, db_name):

        self.DB_data = credentials.db_details[db_name]
        self.connection = pymysql.connect(**self.DB_data)
        self.cursor = self.connection.cursor()
