#!/usr/bin/python
# -*- coding:utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import xlrd
from xml.dom.minidom import Document
import Tkinter,tkFileDialog
import os

class excel_to_xml():

    def __init__(self):
        pass

    def get_filepath(self):
        """
        提供文件选择对话框，从本地选择要转化的excel文件，返回文件路径
        """
        root = Tkinter.Tk()
        root.withdraw()
        filepath = tkFileDialog.askopenfilename('/Users/hudi/PycharmProjects/untitled5/explore/导入文档.xls')
        return filepath


    def read_excel(self,path):
        """
        读取excel内容，返回所有用例值
        """
        # 打开文件
        file = xlrd.open_workbook(path)
        # 获取第一个sheet（按索引）
        sheet1 = file.sheet_by_index(0)

        # 获取行数和列数
        nrows = sheet1.nrows
        ncols = sheet1.ncols

        print nrows,ncols

        # 获取单元格内容
        nclosvalue = []
        for j in range(1,nrows):
            nrowsvalue = []
            for i in range(ncols):
                cellvalue = sheet1.cell(j,i)
                nrowsvalue.append(cellvalue)
                i +=1


            nclosvalue.append(nrowsvalue)
            j += 1

        return nclosvalue


    def to_xml(self):
        """
        处理数据，转化成xml格式，并将文件保存在用例同路径下
        """
        path1 = self.get_filepath()
        doc = Document()  # 创建DOM文档对象

        testcases = doc.createElement('testcases')
        doc.appendChild(testcases)

        excle_results = self.read_excel(path1)
        print(len(excle_results))
        for i in range(len(excle_results)):
            print"第"+str(i+1)+"个用例为：\n"
            print(excle_results[i])

            testcase = doc.createElement('testcase')
            testcase.setAttribute('name', "%s" % excle_results[i][0].value)
            testcases.appendChild(testcase)

            summary = doc.createElement('summary')
            summary_text = doc.createTextNode('%s' % excle_results[i][1].value)
            summary.appendChild(summary_text)
            testcase.appendChild(summary)

            steps = doc.createElement('steps')
            testcase.appendChild(steps)

            step = doc.createElement('step')
            steps.appendChild(step)

            step_number = doc.createElement('step_number')
            step_number_text = doc.createTextNode('1')
            step_number.appendChild(step_number_text)
            step.appendChild(step_number)

            actions = doc.createElement('actions')
            actions_text = doc.createTextNode('%s' % excle_results[i][2].value)
            actions.appendChild(actions_text)
            step.appendChild(actions)

            expectedresults = doc.createElement('expectedresults')
            expectedresults_text = doc.createTextNode('%s' % excle_results[i][3].value)
            expectedresults.appendChild(expectedresults_text)
            step.appendChild(expectedresults)

            i += 1

        # 要生成的xml文件名
        xml_name = path1.strip().split('.')[0] + '.xml'

        # 要生成的xml文件到目录（绝对路径）
        dir = path1.strip().split('/')[-2]
        xml_dir = os.path.join(('%s') % dir,xml_name)

        try:
            f = open(xml_dir,'w')
            doc.writexml(f, indent='\t', newl='\n', addindent='\t', encoding='utf-8')
            f.close()
        except:
            print("您没有选择任何文件！")




change = excel_to_xml()
change.to_xml()