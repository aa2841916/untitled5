# coding:utf-8
import importlib
import os
import sys

from click._compat import raw_input

importlib.reload(sys)
# sys.setdefaultencoding("utf-8")

from explore import easy_excel
class operate():
    def __init__(self, ExcelFileName, SheetName):
        self.excelFile = ExcelFileName + '.xls'
        self.excelSheet = SheetName
        self.temp = easy_excel(self.excelFile)
        self.dic_testlink = {}
        self.row_flag = 3
        self.testsuite = self.temp.getCell(self.excelSheet, 2, 1)
        self.dic_testlink[self.testsuite] = {"node_order": "13", "details": "", "testcase": []}
        self.content = ""
        self.content_list = []

    def xlsx_to_dic(self, SheetName):
        while True:
            # print 'loop1'
            # list_testcase = dic_testlink[testsuite].["testcase"]

            testcase = {"name": "", "node_order": "100", "externalid": "", "version": "1", "summary": "",
                        "preconditions": "", "execution_type": "1", "importance": "3", "steps": [], "keywords": "P1"}
            testcase["name"] = self.temp.getCell(self.excelSheet, self.row_flag, 1)
            testcase["summary"] = self.temp.getCell(self.excelSheet, self.row_flag, 3)
            testcase["preconditions"] = self.temp.getCell(self.excelSheet, self.row_flag, 4)
            execution_type = self.temp.getCell(self.excelSheet, self.row_flag, 7)
            if execution_type == "自动":
                testcase["execution_type"] = 2
            # print self.temp.getCell('Sheet1',self.row_flag,3)
            step_number = 1
            testcase["keywords"] = self.temp.getCell(self.excelSheet, self.row_flag, 2)
            # print testcase["keywords"]
            while True:
                # print 'loop2'
                step = {"step_number": "", "actions": "", "expectedresults": "", "execution_type": ""}
                step["step_number"] = step_number
                step["actions"] = self.temp.getCell(self.excelSheet, self.row_flag, 5)
                step["expectedresults"] = self.temp.getCell(self.excelSheet, self.row_flag, 6)
                testcase["steps"].append(step)
                step_number += 1
                self.row_flag += 1
                if self.temp.getCell(self.excelSheet, self.row_flag, 1) is not None or self.temp.getCell(self.excelSheet, self.row_flag, 5) is None:
                    break
            # print testcase

            self.dic_testlink[self.testsuite]["testcase"].append(testcase)
            # print self.row_flag
            if self.temp.getCell(self.excelSheet, self.row_flag, 5) is None and self.temp.getCell(self.excelSheet, self.row_flag + 1, 5) is None:
                break
        self.temp.close()
        # print self.dic_testlink

    def content_to_xml(self, key, value=None):
        if key == 'step_number' or key == 'execution_type' or key == 'node_order' or key == 'externalid' or key == 'version' or key == 'importance':
            return "<" + str(key) + "><![CDATA[" + str(value) + "]]></" + str(key) + ">"
        elif key == 'actions' or key == 'expectedresults' or key == 'summary' or key == 'preconditions':
            return "<" + str(key) + "><![CDATA[<p> " + str(value) + "</p> ]]></" + str(key) + ">"
        elif key == 'keywords':
            return '<keywords><keyword name="' + str(value) + '"><notes><![CDATA[ aaaa ]]></notes></keyword></keywords>'
        elif key == 'name':
            return '<testcase name="' + str(value) + '">'
        else:
            return '##########'

    def dic_to_xml(self, ExcelFileName, SheetName):
        testcase_list = self.dic_testlink[self.testsuite]["testcase"]
        for testcase in testcase_list:
            for step in testcase["steps"]:
                self.content += "<step>"
                self.content += self.content_to_xml("step_number", step["step_number"])
                self.content += self.content_to_xml("actions", step["actions"])
                self.content += self.content_to_xml("expectedresults", step["expectedresults"])
                self.content += self.content_to_xml("execution_type", step["execution_type"])
                self.content += "</step>"
            self.content = "<steps>" + self.content + "</steps>"
            self.content = self.content_to_xml("importance", testcase["importance"]) + self.content
            self.content = self.content_to_xml("execution_type", testcase["execution_type"]) + self.content
            self.content = self.content_to_xml("preconditions", testcase["preconditions"]) + self.content
            self.content = self.content_to_xml("summary", testcase["summary"]) + self.content
            self.content = self.content_to_xml("version", testcase["version"]) + self.content
            self.content = self.content_to_xml("externalid", testcase["externalid"]) + self.content
            self.content = self.content_to_xml("node_order", testcase["node_order"]) + self.content
            self.content = self.content + self.content_to_xml("keywords", testcase["keywords"])
            self.content = self.content_to_xml("name", testcase["name"]) + self.content
            self.content = self.content + "</testcase>"
            self.content_list.append(self.content)
            self.content = ""
        self.content = "".join(self.content_list)
        self.content = '<testsuite name="' + self.testsuite + '">' + self.content + "</testsuite>"
        self.content = '<?xml version="1.0" encoding="UTF-8"?>' + self.content
        self.write_to_file(ExcelFileName, SheetName)

    def write_to_file(self, ExcelFileName, SheetName):
        xmlFileName = ExcelFileName + '_' + SheetName + '.xml'
        cp = open(xmlFileName, "w")
        cp.write(self.content)
        cp.close()

if __name__ == "__main__":

    fileName = raw_input('enter excel name:')
    sheetName = raw_input('enter sheet name:')
    sheetList = sheetName.split(" ")
    for sheetName in sheetList:
        test = operate(fileName, sheetName)
        test.xlsx_to_dic(sheetName)
        test.dic_to_xml(fileName, sheetName)
    print("Convert success!")
    os.system('pause')