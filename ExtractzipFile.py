# get all zip files from current directory
# unzip each file 1 by 1
import os
from zipfile import ZipFile
from selenium import webdriver
import time
import os
import asq
import win32com.client



class Testcase:
    def __init__(self, name, result, testplanname,tcEndTime):
        self.name = name
        self.result = result
        self.testPlanName = testplanname
        self.tcEndTime = tcEndTime
    def __repr__(self):
        return self


def downloadEmailAttachment():



    Cmd = "del *.zip"
    os.system(Cmd)

    Cmd = "del Report*.html"
    os.system(Cmd)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
    inbox = inbox.Folders['springmvc-reducedprodset']
    messages = inbox.Items
    message = messages.GetFirst()

    while message is not None:
        try:
            print (message)
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile(os.getcwd() + '\\' + str(attachment)) #Saves to the attachment to current folder
            print (attachment)
            message = messages.GetNext()

        except:
            message = messages.GetNext()

def extractZipFile():
    Cmd = "dir /b *.zip"
    fileList = os.popen(Cmd).read()
    fileList = fileList.split('\n')

    for i in fileList:
        if len(i.strip()) > 0:
            with ZipFile(i, 'r') as zip:
                zip.printdir()
                # printing all the contents of the zip file
                # extracting all the files
                print('Extracting all the files now...')
                zip.extractall()


# now

def readTestResults():
    global testCaseList
    testCaseList = []

    driver = webdriver.Chrome()
    Cmd = "dir /b *.html"
    currentDirPath = os.getcwd()
    print("path = " + currentDirPath)
    fileList = os.popen(Cmd).read()
    fileList = fileList.split('\n')

    for i in fileList:
        if len(i.strip()) > 0:
            print("PATH = "+currentDirPath + "\\" + i)
            driver.get(currentDirPath + "\\" + i)
           # testcaseLink = driver.find_element_by_xpath("//*[@id='slide-out']/li[2]/a/i")
           # testcaseLink.click()
            testcaseList = driver.find_elements_by_xpath("//ul[@id='test-collection']/li")
            print("=====================================" + i + "=====================================")
            cnt = 0
            for div in testcaseList:
                cnt = cnt + 1
                xpath = "//*[@id='test-collection']/li[" + str(cnt) + "]/div[1]/span[1]";
                # print("xpath = " + xpath)
                testcaseName = driver.find_element_by_xpath(
                    "//*[@id='test-collection']/li[" + str(cnt) + "]/div[1]/span[1]")
                result = driver.find_element_by_xpath("//*[@id='test-collection']/li[" + str(cnt) + "]/div[1]/span[3]")
                testplanName = driver.find_element_by_xpath(
                    "//*[@id='test-collection']/li[" + str(cnt) + "]/div[2]/div[3]/div/span[2]")

             #   tcEndTime = driver.find_element_by_xpath("//*[@id='test-collection']/li[" + str(cnt) + "]div[2]/div[1]/span[2]")
                tcEndTime = ""
                if testcaseName.is_displayed():
                    print(testplanName.get_attribute("innerHTML") + " - " + testcaseName.text + " " + result.text)
                    testCaseList.append(Testcase(testcaseName.text, result.text, testplanName.get_attribute("innerHTML"), tcEndTime))
                   # duplicateList = asq.query(testCaseList).where(asq.query(testCaseList).count(lambda x: x.testcaseName) > 1).to_list()
   # testCaseList = asq.query(testCaseList).group_by(lambda x: x.name).select(lambda x: x.)
   # testCaseList = asq.query(testCaseList).select(lambda x: Testcase(x.name, x.result, x.testPlanName, x.tcEndTime)).to_list()

    driver.quit()

def createResultFile():
    f = open('result.html','w')
    message ="<html><head><style>.pass { width: 100%; text-align: left; color: green } .fail { text-align: left;color: red;font-weight: bold }</style></head><body><table border='1'><tr><th>TestPlan</th><th>Testcase</th><th>Result</th></tr>"
    for obj in testCaseList:

        if str(obj.result).lower() in "fail":
            message = message + "<tr class='fail'>"
        else:
            message = message + "<tr class='pass'>"
        message = message + "<td>" + str(obj.testPlanName) + "</td>"
        message = message + "<td>" + str(obj.name) + "</td>"
        message = message + "<td>" + str(obj.result) + "</td>"
        message = message + "</tr>"

    message = message + "</body></html>"
    f.write(message)
    f.close()

def main():
    downloadEmailAttachment()
    extractZipFile()
    readTestResults()
    createResultFile()


main()
# Pending

# if zip file is not there skip the extraction
# bind all these into class
# display testplan wise results
# failures should be on top
# display summary
# fetch zip files from outlook folders
# create result in html format
