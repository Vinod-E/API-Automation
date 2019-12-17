import unittest
from collections import OrderedDict
import datetime
import xlwt
import time
import requests
import random
import json
from Config import Api
from scripts.assessment import WebConfig
from selenium import webdriver
from Config.read_excel import *


class Candidate_Transcript(unittest.TestCase):
    @classmethod
    def setUp(cls):
        cls.driver = webdriver.Chrome(WebConfig.CHROME_DRIVER)
        cls.driver.implicitly_wait(30)
        cls.driver.maximize_window()
        cls.driver.get(WebConfig.ONLINE_ASSESSMENT_LOGIN_URL)
        return cls.driver

    def test_Candidate_Transcript(self):
        now = datetime.datetime.now()
        __current_DateTime = now.strftime("%d/%m/%Y")
        answers = ['A', 'B', 'C', 'D']
        # --------------------------------------------------------------------------------------------------------------
        # CSS to differentiate Correct and Wrong data in Excel
        # --------------------------------------------------------------------------------------------------------------
        self.__style0 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold on; pattern: pattern solid, fore-colour gold; border: left thin,right thin,top thin,bottom thin')
        self.__style1 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; border: left thin,right thin,top thin,bottom thin')
        self.__style2 = xlwt.easyxf(
            'font: name Times New Roman, color-index red, bold on; border: left thin,right thin,top thin,bottom thin')
        self.__style3 = xlwt.easyxf(
            'font: name Times New Roman, color-index green, bold on; border: left thin,right thin,top thin,bottom thin')
        self.__style4 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; pattern: pattern solid, fore-colour light_yellow; border: left thin,right thin,top thin,bottom thin')
        self.__style5 = xlwt.easyxf(
            'font: name Times New Roman, color-index black, bold off; pattern: pattern solid, fore-colour yellow; border: left thin,right thin,top thin,bottom thin')
        # mycursor.execute('delete from test_result_infos where testresult_id in (select id from test_results where testuser_id in (select id from test_users where test_id = '+i+' and login_time is not null));')
        # --------------------------------------------------------------------------------------------------------------
        # Read from Excel
        # --------------------------------------------------------------------------------------------------------------
        excel_read_obj.excel_read('/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Input Data/Assessment/ClientQuestionRandomization.xls', 0)
        self.xls_values = excel_read_obj.details
        wb_result = xlwt.Workbook()
        self.ws = wb_result.add_sheet('questionRandomization', cell_overwrite_ok=True)
        col_index = 0
        # -------------------------Need to change file name onle here based on Test level configuration in UI---------------------
        self.file = open("/home/rajeshwar/D Drive/hirepro_automation/API-Automation/Output Data/Assessment/Client_Test_Random_Check.html", "wt")
        # file name need to be changed based on Test level configuration (
        # Client_Test_Random_Check.html or
        # Client_Group_Random_Check.html or
        # Client_Section_Random_Check.html))
        self.file.write("""<html>
                       <head>
                       <title>Automation Results</title>
                       <style>
                       h1 {
                           color: #0e8eab;
                           text-align: left;
                           font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
                       }
                       .div-h1 {
                           position: absolute;
                           overflow: hidden;
                           top: 0;
                           width: auto;
                           height: 100px;
                           text-align: center;
                       }
                       .div-overalldata {
                           position: absolute;
                           top: 80px;
                           width: 600px;
                           height: auto;
                           text-align: left;
                           font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
                       }
                       .label {
                           color: #0e8eab;
                           font-family: Arial;
                           font-size: 14pt;
                           font-weight: bold;
                       }
                       .value {
                           color: black;
                           font-family: Arial;
                           font-size: 14pt;
                       }
                       .valuePass {
                           color: green;
                           font-family: Arial;
                           animation: blinkingTextPass 0.8s infinite;
                           font-weight: bold;
                           font-size: 20pt;
                       }       
                       @keyframes blinkingTextPass{
                           0%{     color: green; font-size: 0pt;  }
                           50%{    color: lightgreen; }
                           100%{   color: green; font-size: 14pt; } 
                       }
                       .valueFail {
                           color: red;
                           font-family: Arial;
                           animation: blinkingTextFail 0.8s infinite;
                           font-weight: bold;
                           font-size: 20px;
                       }
                       @keyframes blinkingTextFail{
                           0%{     color: red; font-size: 0pt;   }
                           50%{    color: orange; }
                           100%{   color: red; font-size: 14pt;  }
                       }
                       .zui-table {
                           border: none;
                           border-right: solid 1px #DDEFEF;
                           border-collapse: separate;
                           border-spacing: 0;
                           font: normal 13px Arial, sans-serif;
                           width: 100%
                       }
                       .zui-table thead th {
                           border-left: solid 1px white;
                           border-bottom: solid 1px #DDEFEF;
                           background-color: #0e8eab;
                           color: white;
                           padding: 10px;
                           text-align: left;
                           white-space: nowrap;
                       }
                       .zui-table tbody td {
                           border-left: solid 1px #DDEFEF;
                           border-right: solid 1px #DDEFEF;
                           border-bottom: solid 1px #DDEFEF;
                           padding: 10px;
                           white-space: nowrap;
                       }
                       .td-pass {
                           color: green;
                       }
                       .td-fail {
                           color: red;
                           font-weight: bold;
                       }
                       tr:nth-child(even){background-color: #f2f2f2;}      
                       tr:hover {background-color: #ddd; border-collapse: collapse;}
                       @media all{
                           table tr th:nth-child(2),
                           table tr td:nth-child(2),
                           table tr th:nth-child(5),
                           table tr td:nth-child(5),
                           table tr th:nth-child(6),
                           table tr td:nth-child(6){
                               display: none;
                           }
                       }
                       .zui-wrapper {
                           position: relative;
                           top: 180px;
                           width: 100%;
                           height: 100%;
                       }
                       .zui-scroller {
                           margin-left: 141px;
                           overflow-x: scroll;
                           overflow-y: visible;
                           padding-bottom: 5px;
                       }
                       .zui-table .zui-sticky-col {
                           border-left: solid 1px #DDEFEF;
                           border-right: solid 1px #DDEFEF;
                           left: 0;
                           position: absolute;
                           top: auto;
                           width: 120px;
                       }
                       .zui-table .zui-sticky-col-pass {
                           border-left: solid 1px #DDEFEF;
                           border-right: solid 1px #DDEFEF;
                           left: 0;
                           position: absolute;
                           top: auto;
                           width: 120px;
                           color:green;
                           font-weight: bold;
                       }
                       .zui-table .zui-sticky-col-fail {
                           border-left: solid 1px #DDEFEF;
                           border-right: solid 1px #DDEFEF;
                           left: 0;
                           position: absolute;
                           top: auto;
                           width: 120px;
                           color:red;
                           font-weight: bold;
                       }
                       .zui-table1 {
                           border: none;
                           border-right: solid 1px #DDEFEF;
                           border-collapse: separate;
                           border-spacing: 0;
                           font: normal 13px Arial, sans-serif;
                           width: 100%
                       }
                       .zui-table1 thead th {
                           border-left: solid 1px white;
                           border-bottom: solid 1px #DDEFEF;
                           background-color: #0e8eab;
                           color: white;
                           padding: 10px;
                           text-align: left;
                           white-space: nowrap;
                       }
                       .zui-table1 tbody td {
                           border-left: solid 1px #DDEFEF;
                           border-right: solid 1px #DDEFEF;
                           border-bottom: solid 1px #DDEFEF;
                           padding: 10px;
                           white-space: nowrap;
                       }
                       .zui-wrapper1 {
                           width: 50%;
                           position: absolute;
                           top: 2%;
                           left: 49%;
                       }
                       </style>
                       <div class="div-h2">

                           <h1>Client Side Test Level Randomization Report</h1></div>
                       </head>
                       <body style="overflow: hidden;">
                       <div class="zui-wrapper">
                       <div class="zui-scroller"><table class="zui-table"><thead><tr>""")
        # <h1>Client Side Test Level Randomization Report</h1></div> Need to change Title as mentioned in file name
        for xls_headers in excel_read_obj.headers_available_in_excel:
            self.ws.write(0, col_index, xls_headers, self.__style0)
            self.file.write(("""<th>""" + str(xls_headers) + """</th>"""))
            col_index += 1
        self.file.write("""<th class="zui-sticky-col">Status</th>""")
        self.file.write("""</tr></thead><tbody>""")
        self.cellPos = 0
        self.rownum = 1
        overall_status = []
        for login_details in self.xls_values:
            candidateId = int(login_details.get('Candidate Id'))