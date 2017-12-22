# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\CIM001\Documents\pyProject\mainwindow.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

import sys
import json
import time
import visual_data_func as vf
from PyQt5 import QtCore, QtGui, QtWidgets

now_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(569, 269)
        self.centralWidget = QtWidgets.QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.label = QtWidgets.QLabel(self.centralWidget)
        self.label.setGeometry(QtCore.QRect(30, 20, 200, 16))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralWidget)
        self.label_2.setGeometry(QtCore.QRect(30, 70, 200, 16))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralWidget)
        self.label_3.setGeometry(QtCore.QRect(30, 120, 200, 16))
        self.label_3.setObjectName("label_3")
        self.block_list_text = QtWidgets.QLineEdit(self.centralWidget)
        self.block_list_text.setGeometry(QtCore.QRect(30, 40, 113, 20))
        self.block_list_text.setObjectName("block_list_text")
        self.matid_list_text = QtWidgets.QLineEdit(self.centralWidget)
        self.matid_list_text.setGeometry(QtCore.QRect(30, 90, 113, 20))
        self.matid_list_text.setObjectName("matid_list_text")
        self.bottom_time_text = QtWidgets.QLineEdit(self.centralWidget)
        self.bottom_time_text.setGeometry(QtCore.QRect(30, 140, 81, 20))
        self.bottom_time_text.setObjectName("bottom_time_text")
        self.top_time_text = QtWidgets.QLineEdit(self.centralWidget)
        self.top_time_text.setGeometry(QtCore.QRect(120, 140, 81, 20))
        self.top_time_text.setObjectName("top_time_text")
        self.excel_textBrowser = QtWidgets.QTextBrowser(self.centralWidget)
        self.excel_textBrowser.setGeometry(QtCore.QRect(225, 10, 321, 151))
        self.excel_textBrowser.setObjectName("excel_textBrowser")
        self.change_config_push_button = QtWidgets.QPushButton(self.centralWidget)
        self.change_config_push_button.setGeometry(QtCore.QRect(40, 180, 101, 31))
        self.change_config_push_button.setObjectName("change_config_push_button")
        self.create_excel_push_button = QtWidgets.QPushButton(self.centralWidget)
        self.create_excel_push_button.setGeometry(QtCore.QRect(210, 180, 91, 31))
        self.create_excel_push_button.setObjectName("create_excel_push_button")
        MainWindow.setCentralWidget(self.centralWidget)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 569, 23))
        self.menuBar.setObjectName("menuBar")
        MainWindow.setMenuBar(self.menuBar)
        self.mainToolBar = QtWidgets.QToolBar(MainWindow)
        self.mainToolBar.setObjectName("mainToolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)

        # button action
        self.change_config_push_button.clicked.connect(self.change_config)
        self.create_excel_push_button.clicked.connect(self.create_excel)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "BLOCK_ID 用英文逗号隔开"))
        self.label_2.setText(_translate("MainWindow", "MAT_ID 用英文逗号隔开"))
        self.label_3.setText(_translate("MainWindow", "TIME 例:2017-10-01 14"))
        self.change_config_push_button.setText(_translate("MainWindow", "修改配置信息"))
        self.create_excel_push_button.setText(_translate("MainWindow", "创建Excel"))

    def message_box(self, warning_msg):
        qtm = QtWidgets.QMessageBox
        msg_box = qtm(qtm.Warning, "Error", warning_msg)
        msg_box.exec_()
        return

    def change_config(self):
        json_dict = {}
        if ((self.block_list_text.text() == "") and (self.matid_list_text.text() == "")) or ((self.block_list_text.text() != "") and (self.matid_list_text.text() != "")):
            self.message_box('BLOCK_LIST OR MAT_LIST IS ILLEGAL')

        elif self.block_list_text.text() != "":
            json_dict['BLOCK'] = self.block_list_text.text().split(',')
        else:
            json_dict['MAT'] = self.matid_list_text.text().split(',')

        if self.top_time_text.text() != "" and self.bottom_time_text.text() != "":
            json_dict['TIME'] = [self.bottom_time_text.text(), self.top_time_text.text()]
        else:
            self.message_box('TIME IS ILLEGAL')

        print(json_dict)
        if json_dict != {}:
            self.excel_textBrowser.append('%s: config info: %s\n' % (now_time, json_dict))
            # update config file
            with open('tqs_excel_config.json', 'w') as json_file:
                json_file.write(json.dumps(json_dict))
            self.excel_textBrowser.append('%s: tqs_excel_config.json is updated sucessfully!\n' % now_time)
        return

    def create_excel(self):
        connection = vf.connect2DB()
        # plot_tqs_summary_data_df(connection, spec_id_dict)
        # plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
        try:
            vf.tqs_summary_data_excel(connection, config_path='tqs_excel_config.json')
            self.excel_textBrowser.append('%s: Excel generated successfully!\n' % now_time)
        finally:
            self.excel_textBrowser.append('%s: Error occurred when creating excel! Please call 18116398262 for help.\n' % now_time)
            self.message_box('Error occurred when creating excel')

        return




app = QtWidgets.QApplication(sys.argv)
hello = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(hello)
hello.show()
app.exec_()

