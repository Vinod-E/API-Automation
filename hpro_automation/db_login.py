from hpro_automation.identity import credentials
import pymysql
from hpro_automation.api import *


class DBConnection(object):

    def __init__(self):
        super(DBConnection, self).__init__()
        self.DB_data = ''
        self.connection = ''
        self.cursor = ''

    def db_connection(self):

        if login_server == 'amsin':
            self.DB_data = credentials.db_details['crpo_amsin']
            self.connection = pymysql.connect(**self.DB_data)
            self.cursor = self.connection.cursor()
        else:
            self.DB_data = credentials.db_details['crpo_ams']
            self.connection = pymysql.connect(**self.DB_data)
            self.cursor = self.connection.cursor()

    def db_connection_tenant(self):

        if login_server == 'amsin':
            self.DB_data = credentials.db_details['tenant_amsin']
            self.connection = pymysql.connect(**self.DB_data)
            self.cursor = self.connection.cursor()
        else:
            self.DB_data = credentials.db_details['tenant_ams']
            self.connection = pymysql.connect(**self.DB_data)
            self.cursor = self.connection.cursor()
