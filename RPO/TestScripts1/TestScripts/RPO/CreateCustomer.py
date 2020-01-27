import unittest
import time

import datetime
from sqlite3 import Date

import xlrd
import xlwt
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from urlparse import urlparse
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


class createCustomer(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new browser session """
        inst.driver = webdriver.Chrome("/home/rajeshwar/Downloads/chromedriver")
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get("http://10.0.3.41/rpo")
        print ('\nEntered URL in browser')
        inst.driver.title
        return inst.driver

    def test_createCandidate(self):
        print ('User reached on project selection screen')
        self.driver.find_element_by_name("clientEmail").send_keys("0003")
        self.driver.find_element_by_name("new-password").send_keys("Ujjwal@123")
        self.driver.find_element_by_xpath("//div[2]/div/div[1]/div/div/div[4]/div[1]/div[1]/div/button").click()
        time.sleep(3)
        self.driver.find_element_by_xpath("//li[4]/a").click()
        time.sleep(2)
        self.driver.find_element_by_xpath("//p[2]").click()
        time.sleep(2)
        CustomerName = self.driver.find_element_by_name("name")
        CustomerName.clear()
        CustomerName.send_keys("Customer_Script_One")
        time.sleep(1)

        Alias = self.driver.find_element_by_xpath("//div[2]/md-input-container/input")
        Alias.clear()
        Alias.send_keys("CSO")
        time.sleep(1)

        ContractAttachment = self.driver.find_element_by_css_selector("input[type='file']")
        ContractAttachment.send_keys("/home/rajeshwar/Downloads/CDC_UP_Test_Plan_Template.doc")
        time.sleep(1)

        ContractSnapshotAttachment = self.driver.find_element_by_xpath("(//input[@type='file'])[2]")
        ContractSnapshotAttachment.send_keys("/home/rajeshwar/Pictures/Wallpapers/tumblr_inline_ntmcp6MXoB1s78p8g_500.jpg")
        time.sleep(1)

        Industry = self.driver.find_element_by_xpath("//div[3]/md-input-container/md-select")
        Industry.send_keys("IT")
        time.sleep(1)

        SubDomain = self.driver.find_element_by_xpath("//div[4]/md-input-container/md-select")
        SubDomain.send_keys("Accounts/Finance")
        time.sleep(1)

        PrimaryLocation = self.driver.find_element_by_xpath("//div[5]/md-input-container/md-select")
        PrimaryLocation.send_keys("Aligarh")
        time.sleep(1)

        OfficeLocation = self.driver.find_element_by_xpath("//div[6]/md-input-container/md-select")
        OfficeLocation.send_keys("Ahmedabad")
        time.sleep(1)

        Status = self.driver.find_element_by_xpath("//div[7]/md-input-container/md-select")
        Status.send_keys("Active")
        time.sleep(1)

        OwnershipTypeDual = self.driver.find_element_by_xpath("//md-radio-button/div/div")
        OwnershipTypeDual.click()
        time.sleep(1)

        HAccountManager = self.driver.find_element_by_xpath("//div[2]/div/md-input-container/md-select")
        HAccountManager.send_keys("A G Sakthi Priya ")
        time.sleep(1)

        HAccountManagerValue = self.driver.find_element_by_xpath("//div[6]/md-select-menu/md-content/md-optgroup/md-option[3]/div")
        HAccountManagerValue.click()
        time.sleep(2)

        VAccountManager = self.driver.find_element_by_xpath("//div[2]/md-input-container/md-select")
        VAccountManager.send_keys("A G Sakthi Priya ")
        time.sleep(1)

        ExecutiveSponsor = self.driver.find_element_by_xpath("//div[2]/div[3]/md-input-container/md-select")
        ExecutiveSponsor.send_keys("A G Sakthi Priya ")
        time.sleep(1)

        ExecutiveSponsorValue = self.driver.find_element_by_xpath("//div[8]/md-select-menu/md-content/md-optgroup/md-option[4]/div")
        ExecutiveSponsorValue.click()
        time.sleep(2)

        CustomerEntityName = self.driver.find_element_by_name("contractname")
        CustomerEntityName.clear()
        CustomerEntityName.send_keys("ABCDEFGH")
        time.sleep(1)

        Entity = self.driver.find_element_by_xpath("//div[7]/div/div[2]/md-input-container/md-select")
        Entity.send_keys("HirePro")
        time.sleep(1)

        StartDate = self.driver.find_element_by_xpath("//div/input")
        StartDate.clear()
        StartDate.send_keys("1/2/1999")
        time.sleep(1)

        EndDateType = self.driver.find_element_by_xpath("//div[7]/div/div[4]/md-input-container/md-select")
        EndDateType.send_keys("End Date Specified")
        time.sleep(3)

        EndDateTypeValue =self.driver.find_element_by_xpath("//div[9]/md-select-menu/md-content/md-option[2]")
        EndDateTypeValue.click()
        time.sleep(2)

        EndDate = self.driver.find_element_by_xpath("//div[5]/md-datepicker/div/input")
        EndDate.clear()
        EndDate.send_keys("1/2/2015")
        time.sleep(1)

        StatusType = self.driver.find_element_by_xpath("//div[7]/div/div[6]/md-input-container/md-select")
        StatusType.send_keys("Contract Signed")
        time.sleep(1)

        StatusTypeValue = self.driver.find_element_by_xpath("//div[10]/md-select-menu/md-content/md-option[4]/div")
        StatusTypeValue.click()
        time.sleep(2)

        ContractStatus = self.driver.find_element_by_xpath("//div[7]/div/div[7]/md-input-container/md-select")
        ContractStatus.send_keys("Both parties signed")
        time.sleep(1)

        ContractStatusValue = self.driver.find_element_by_xpath("//div[11]/md-select-menu/md-content/md-option[2]/div")
        ContractStatusValue.click()
        time.sleep(2)

        ServiceLine = self.driver.find_element_by_xpath("//div[7]/div/div[8]/md-input-container/md-select")
        ServiceLine.send_keys("Consulting")
        time.sleep(1)

        DocumentType = self.driver.find_element_by_xpath("//div[7]/div/div[9]/md-input-container/md-select")
        DocumentType.send_keys("NDA")
        time.sleep(1)

        Remark = self.driver.find_element_by_xpath("//div[10]/md-input-container/input")
        Remark.clear()
        Remark.send_keys("Test Remark")
        time.sleep(1)

        # ClickAnywhereOnScreen = self.driver.find_element_by_xpath("//md-backdrop")
        # ClickAnywhereOnScreen.click()

        AddContract = self.driver.find_element_by_xpath("(//button[@type='button'])[8]")
        AddContract.click()
        time.sleep(2)

        SaveContract = self.driver.find_element_by_xpath("(//button[@type='button'])[8]")
        SaveContract.click()


        time.sleep(10)



    @classmethod
    def tearDown(inst):
        # close the browser window
        inst.driver.quit()


if __name__ == '__main__':
    unittest.main()