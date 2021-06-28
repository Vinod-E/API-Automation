import re
import json
import pandas as pd


class HistoryDataRead:

    def __init__(self, path):
        self.path = path
        self.history_data_AMSIN_NON_EU_dict = []
        self.history_data_AMSIN_EU_dict = []
        self.history_data_LIVE_NON_EU_dict = []
        self.history_data_LIVE_EU_dict = []

        self.pass_cases = []
        self.time_taken = []

        self.graph_sprint_names = []

        self.amsin_pass_cases = []
        self.amsin_execute_time = []

        self.beta_pass_cases = []
        self.beta_execute_time = []

        self.ams_pass_cases = []
        self.ams_execute_time = []

        self.india_pass_cases = []
        self.india_execute_time = []

        self.version_number = ''

    def read_excel_data_dict(self):
        print('**----->> History Data reading from History Excel')

        """ ======== Open and Read Excel to get the sheet data ==========="""
        reader = pd.read_excel(self.path, engine='openpyxl', index_col=0)
        self.history_data_AMSIN_NON_EU_dict = json.loads(reader.to_json(orient='records'))
        # print('Excel Sheet to JSON:\n', self.history_data_AMSIN_dict)

        reader = pd.read_excel(self.path, engine='openpyxl', index_col=1)
        self.history_data_AMSIN_EU_dict = json.loads(reader.to_json(orient='records'))
        # print('Excel Sheet to JSON:\n', self.history_data_BETA_dict)

        reader = pd.read_excel(self.path, engine='openpyxl', index_col=2)
        self.history_data_LIVE_NON_EU_dict = json.loads(reader.to_json(orient='records'))
        # print('Excel Sheet to JSON:\n', self.history_data_AMS_dict)

        reader = pd.read_excel(self.path, engine='openpyxl', index_col=3)
        self.history_data_LIVE_EU_dict = json.loads(reader.to_json(orient='records'))
        # print('Excel Sheet to JSON:\n', self.history_data_INDIA_dict)

    def amsin_non_eu_dict(self, current_version):
        self.last_five_history_data(current_version, self.history_data_AMSIN_EU_dict)
        self.amsin_pass_cases = self.pass_cases
        self.amsin_execute_time = self.time_taken

    def amsin_eu_dict(self, current_version):
        self.last_five_history_data(current_version, self.history_data_AMSIN_EU_dict)
        self.beta_pass_cases = self.pass_cases
        self.beta_execute_time = self.time_taken

    def ams_non_eu_dict(self, current_version):
        self.last_five_history_data(current_version, self.history_data_LIVE_NON_EU_dict)
        self.ams_pass_cases = self.pass_cases
        self.ams_execute_time = self.time_taken

    def ams_eu_dict(self, current_version):
        self.last_five_history_data(current_version, self.history_data_LIVE_EU_dict)
        self.india_pass_cases = self.pass_cases
        self.india_execute_time = self.time_taken

    def last_five_history_data(self, current_sprint_number, dicts):
        sort_sprint_names = []
        self.pass_cases = []
        self.time_taken = []

        self.version_number = re.search(r'\d+', current_sprint_number).group()
        number = self.version_number
        attempts = 0
        while attempts < 5:
            sort_sprint_names.append('Sprint_{}'.format(int(number)-1))
            number = int(number) - 1
            attempts += 1
        self.graph_sprint_names = [i for i in sorted(sort_sprint_names)]
        # print('Sprint Names:: ', self.graph_sprint_names)

        self.pass_cases = []
        self.time_taken = []

        all_sprint_names = [x['Run Number'] for x in dicts]

        if len(dicts) > 5:
            if current_sprint_number in all_sprint_names:
                index = all_sprint_names.index(current_sprint_number)
                last_5_data = dicts[index - 5:index]
                # print('Dict_IF :: ', last_5_data)
                for k in last_5_data:
                    self.pass_cases.append(k['Pass Cases'])
                    self.time_taken.append(k['Time Taken'])
            else:
                last_5_data = dicts[-5:]
                # print('Dict_ELSE :: ', last_5_data)
                for k in last_5_data:
                    self.pass_cases.append(k['Pass Cases'])
                    self.time_taken.append(k['Time Taken'])
        else:
            print('***----------->>> History Data excel should have more than 5 rows')
