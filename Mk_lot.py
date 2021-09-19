#증착기 Lot폴더 만들기

import csv
import sys
import os
import shutil
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import uic
import tkinter
from tkinter import filedialog
import openpyxl
import win32com.client as win32

set_dic = {}

with open('source_target_path.txt','r') as f:
    reader = csv.reader(f)
    set_dic = {rows[0]: rows[1].lstrip() for rows in reader}

#UI파일 연결
form_class = uic.loadUiType("MK_ui.ui")[0]

#화면을 띄우는데 사용되는 Class 선언
#QMainWindow Class를 상속
class WindowClass(QMainWindow, form_class) :

    def __init__(self) :
        super().__init__()              #기반 클래스의 생성자 실행 : QMainWindow의 생성자 호출
        self.setupUi(self)

        global set_dic

        self.line_source.setText(set_dic['Source'])
        self.line_target.setText(set_dic['Target'])

        self.push_source.clicked.connect(self.mk_path_s)
        self.push_target.clicked.connect(self.mk_path_t)
        self.push_setting.clicked.connect(self.set_open)
        self.push_run.clicked.connect(self.mk_folder)

#######################################경로 지정
    def mk_path_s(self):
        self.mk_path('Source')

    def mk_path_t(self):
        self.mk_path('Target')

    def mk_path(self,s):
        root = tkinter.Tk()
        root.withdraw

        dir_path = filedialog.askdirectory(parent=root, initialdir="/", title='경로 선택')
        root.destroy()

        if len(dir_path) == 0:
            pass
        else:
            if s == 'Source' : self.line_source.setText(dir_path)
            if s == 'Target' : self.line_target.setText(dir_path)
            set_dic[s] = dir_path

            with open('source_target_path.txt', 'w', newline='') as f:
                writer = csv.writer(f)
                for k, v in set_dic.items():
                    writer.writerow([k, v])
########################################경로 지정 끝

    def set_open(self):
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.Workbooks.Open('Device_set.xlsx')

####################################Lot 폴더 생성 시작
    def mk_folder(self):
        self.f_flag = 0
        self.L_flag = 0

        self.f_list = os.listdir(self.line_source.text())
        #Interlock
        for file in self.f_list:
            if self.line_device.text() in file:
                self.f_flag = 1
            if '복사용 수명 매크로' in file:
                self.L_flag = 1

        if self.f_flag != 1:
            QMessageBox.warning(self, "IVL Interlock", self.line_device.text() + " IVL Source 파일이 없습니다.")
            return

        if self.L_flag != 1:
            QMessageBox.warning(self, "LT Interlock", "수명 Source 파일이 없습니다.")
            return


        #1. 전체 폴더 생성
        self.c_folder(self.line_target.text() + '/' + self.line_folder.text())

        #2. 하위 폴더 생성
        self.c_folder(self.line_target.text() + '/' + self.line_folder.text() + '/' + 'IVL')
        self.c_folder(self.line_target.text() + '/' + self.line_folder.text() + '/' + '수명')

        #3. IVL 파일 / 폴더
        self.c_folder(self.line_target.text() + '/' + self.line_folder.text() + '/' + 'IVL' + '/' + str(self.line_folder)[:6])
        shutil.copy(self.line_source.text() + '/' + self.line_device.text() + ' IVL.xlsm', self.line_target.text() + '/' + self.line_folder.text() + '/' + 'IVL')
        os.rename(self.line_target.text() + '/' + self.line_folder.text() + '/' + 'IVL' + '/' + self.line_device.text() + ' IVL.xlsm'
                  , self.line_target.text() + '/' + self.line_folder.text() + '/' + 'IVL' + '/' + self.line_folder.text() + '.xlsm')

        #4. 수명 파일 생성 및 수명 입력
        shutil.copy(self.line_source.text() + '/' + '복사용 수명 매크로.xlsm', self.line_target.text() + '/' + self.line_folder.text() + '/' + '수명')
        os.rename(self.line_source.text() + '/' + '복사용 수명 매크로.xlsm', self.line_target.text() + '/' + self.line_folder.text() + '/' + '수명' + '복사용 수명 매크로.xlsm'
                  , self.line_source.text() + '/' + '복사용 수명 매크로.xlsm', self.line_target.text() + '/' + self.line_folder.text() + '/' + '수명' + self.line_folder.text() + ' - 수명.xlsm')

        #수명 매크로
        #excel = win32.Dispatch("Excel.Application")
        #excel.Visible = True





    def c_folder(self, dir):
        if not os.path.exists(dir):
            os.makedirs(dir)
        else:
            QMessageBox.warning(self, "Path Interlock", "이미 존재하는 폴더입니다.")
            return


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()