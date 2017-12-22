# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\CIM001\Documents\complete_mainwindow.ui'
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
        MainWindow.resize(588, 268)
        self.centralWidget = QtWidgets.QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralWidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 571, 201))
        self.tabWidget.setObjectName("tabWidget")
        self.excel_tab = QtWidgets.QWidget()
        self.excel_tab.setObjectName("excel_tab")
        self.label = QtWidgets.QLabel(self.excel_tab)
        self.label.setGeometry(QtCore.QRect(20, 20, 151, 16))
        self.label.setObjectName("label")
        self.excel_block_list_text = QtWidgets.QLineEdit(self.excel_tab)
        self.excel_block_list_text.setGeometry(QtCore.QRect(20, 40, 113, 20))
        self.excel_block_list_text.setObjectName("excel_block_list_text")
        self.excel_bottom_time_text = QtWidgets.QLineEdit(self.excel_tab)
        self.excel_bottom_time_text.setGeometry(QtCore.QRect(20, 140, 81, 20))
        self.excel_bottom_time_text.setObjectName("excel_bottom_time_text")
        self.label_2 = QtWidgets.QLabel(self.excel_tab)
        self.label_2.setGeometry(QtCore.QRect(20, 70, 141, 16))
        self.label_2.setObjectName("label_2")
        self.excel_matid_list_text = QtWidgets.QLineEdit(self.excel_tab)
        self.excel_matid_list_text.setGeometry(QtCore.QRect(20, 90, 113, 20))
        self.excel_matid_list_text.setObjectName("excel_matid_list_text")
        self.create_excel_push_button = QtWidgets.QPushButton(self.excel_tab)
        self.create_excel_push_button.setGeometry(QtCore.QRect(460, 90, 101, 31))
        self.create_excel_push_button.setObjectName("create_excel_push_button")
        self.label_3 = QtWidgets.QLabel(self.excel_tab)
        self.label_3.setGeometry(QtCore.QRect(20, 120, 171, 16))
        self.label_3.setObjectName("label_3")
        self.excel_top_time_text = QtWidgets.QLineEdit(self.excel_tab)
        self.excel_top_time_text.setGeometry(QtCore.QRect(110, 140, 81, 20))
        self.excel_top_time_text.setObjectName("excel_top_time_text")
        self.excel_textBrowser = QtWidgets.QTextBrowser(self.excel_tab)
        self.excel_textBrowser.setGeometry(QtCore.QRect(200, 10, 251, 151))
        self.excel_textBrowser.setObjectName("excel_textBrowser")
        self.excel_change_config_push_button = QtWidgets.QPushButton(self.excel_tab)
        self.excel_change_config_push_button.setGeometry(QtCore.QRect(460, 40, 101, 31))
        self.excel_change_config_push_button.setObjectName("excel_change_config_push_button")
        self.tabWidget.addTab(self.excel_tab, "")
        self.bar_tab = QtWidgets.QWidget()
        self.bar_tab.setObjectName("bar_tab")
        self.label_4 = QtWidgets.QLabel(self.bar_tab)
        self.label_4.setGeometry(QtCore.QRect(20, 10, 151, 16))
        self.label_4.setObjectName("label_4")
        self.bar_mat_list_text = QtWidgets.QLineEdit(self.bar_tab)
        self.bar_mat_list_text.setGeometry(QtCore.QRect(20, 30, 113, 20))
        self.bar_mat_list_text.setObjectName("bar_mat_list_text")
        self.label_5 = QtWidgets.QLabel(self.bar_tab)
        self.label_5.setGeometry(QtCore.QRect(20, 70, 171, 16))
        self.label_5.setObjectName("label_5")
        self.bar_bottom_time_text = QtWidgets.QLineEdit(self.bar_tab)
        self.bar_bottom_time_text.setGeometry(QtCore.QRect(20, 90, 81, 20))
        self.bar_bottom_time_text.setObjectName("bar_bottom_time_text")
        self.bar_top_time_text = QtWidgets.QLineEdit(self.bar_tab)
        self.bar_top_time_text.setGeometry(QtCore.QRect(110, 90, 81, 20))
        self.bar_top_time_text.setObjectName("bar_top_time_text")
        self.bar_textBrowser = QtWidgets.QTextBrowser(self.bar_tab)
        self.bar_textBrowser.setGeometry(QtCore.QRect(205, 10, 351, 151))
        self.bar_textBrowser.setObjectName("bar_textBrowser")
        self.bar_config_change_pushButton = QtWidgets.QPushButton(self.bar_tab)
        self.bar_config_change_pushButton.setGeometry(QtCore.QRect(20, 130, 81, 31))
        self.bar_config_change_pushButton.setObjectName("bar_config_change_pushButton")
        self.bar_create_pushButton = QtWidgets.QPushButton(self.bar_tab)
        self.bar_create_pushButton.setGeometry(QtCore.QRect(110, 130, 81, 31))
        self.bar_create_pushButton.setObjectName("bar_create_pushButton")
        self.tabWidget.addTab(self.bar_tab, "")
        self.box_tab = QtWidgets.QWidget()
        self.box_tab.setObjectName("box_tab")
        self.label_6 = QtWidgets.QLabel(self.box_tab)
        self.label_6.setGeometry(QtCore.QRect(20, 10, 101, 16))
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.box_tab)
        self.label_7.setGeometry(QtCore.QRect(20, 60, 91, 16))
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(self.box_tab)
        self.label_8.setGeometry(QtCore.QRect(20, 110, 91, 16))
        self.label_8.setObjectName("label_8")
        self.box_mat_text = QtWidgets.QLineEdit(self.box_tab)
        self.box_mat_text.setGeometry(QtCore.QRect(20, 30, 113, 20))
        self.box_mat_text.setObjectName("box_mat_text")
        self.box_block_text = QtWidgets.QLineEdit(self.box_tab)
        self.box_block_text.setGeometry(QtCore.QRect(20, 80, 113, 20))
        self.box_block_text.setObjectName("box_block_text")
        self.box_bottom_time_text = QtWidgets.QLineEdit(self.box_tab)
        self.box_bottom_time_text.setGeometry(QtCore.QRect(20, 130, 101, 20))
        self.box_bottom_time_text.setObjectName("box_bottom_time_text")
        self.box_top_time_text = QtWidgets.QLineEdit(self.box_tab)
        self.box_top_time_text.setGeometry(QtCore.QRect(140, 130, 101, 20))
        self.box_top_time_text.setObjectName("box_top_time_text")
        self.box_textBrowser = QtWidgets.QTextBrowser(self.box_tab)
        self.box_textBrowser.setGeometry(QtCore.QRect(260, 10, 201, 151))
        self.box_textBrowser.setObjectName("box_textBrowser")
        self.box_change_config_pushButton = QtWidgets.QPushButton(self.box_tab)
        self.box_change_config_pushButton.setGeometry(QtCore.QRect(470, 40, 81, 31))
        self.box_change_config_pushButton.setObjectName("box_change_config_pushButton")
        self.box_create_pushButton = QtWidgets.QPushButton(self.box_tab)
        self.box_create_pushButton.setGeometry(QtCore.QRect(470, 100, 81, 31))
        self.box_create_pushButton.setObjectName("box_create_pushButton")
        self.tabWidget.addTab(self.box_tab, "")
        MainWindow.setCentralWidget(self.centralWidget)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 588, 23))
        self.menuBar.setObjectName("menuBar")
        MainWindow.setMenuBar(self.menuBar)
        self.mainToolBar = QtWidgets.QToolBar(MainWindow)
        self.mainToolBar.setObjectName("mainToolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)

        self.excel_change_config_push_button.clicked.connect(self.excel_change_config)
        self.bar_config_change_pushButton.clicked.connect(self.bar_change_config)
        self.box_change_config_pushButton.clicked.connect(self.box_change_config)
        self.create_excel_push_button.clicked.connect(self.create_excel)
        self.bar_create_pushButton.clicked.connect(self.create_bar)
        self.box_create_pushButton.clicked.connect(self.create_box)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "BLOCK_ID LIST"))
        self.excel_bottom_time_text.setText(_translate("MainWindow", "2017-10-1 14"))
        self.label_2.setText(_translate("MainWindow", "MAT_ID LIST"))
        self.create_excel_push_button.setText(_translate("MainWindow", "创建Excel"))
        self.label_3.setText(_translate("MainWindow", "TIME RANGE 例:2017-10-01 17"))
        self.excel_top_time_text.setText(_translate("MainWindow", "2017-11-28 18"))
        self.excel_change_config_push_button.setText(_translate("MainWindow", "修改配置信息"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.excel_tab), _translate("MainWindow", "Excel"))
        self.label_4.setText(_translate("MainWindow", "MAT_ID_LIST"))
        self.label_5.setText(_translate("MainWindow", "TIME RANGE 例:2017-10-01 17"))
        self.bar_config_change_pushButton.setText(_translate("MainWindow", "修改配置"))
        self.bar_create_pushButton.setText(_translate("MainWindow", "创建柱状图"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.bar_tab), _translate("MainWindow", "Bar"))
        self.label_6.setText(_translate("MainWindow", "MAT_ID LIST"))
        self.label_7.setText(_translate("MainWindow", "BLOCK_ID LIST"))
        self.label_8.setText(_translate("MainWindow", "TIME RANGE"))
        self.box_change_config_pushButton.setText(_translate("MainWindow", "修改配置"))
        self.box_create_pushButton.setText(_translate("MainWindow", "创建BOX"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.box_tab), _translate("MainWindow", "Box"))



    def message_box(self, warning_msg):
        qtm = QtWidgets.QMessageBox
        msg_box = qtm(qtm.Warning, "Error", warning_msg)
        msg_box.exec_()
        return

    def excel_change_config(self):
        with open('tqs_excel_config.json') as json_file:
            json_dict = json.load(json_file)
        if ((self.excel_block_list_text.text() == "") and (self.excel_matid_list_text.text() == "")) or ((self.excel_block_list_text.text() != "") and (self.excel_matid_list_text.text() != "")):
            self.message_box('BLOCK_LIST OR MAT_LIST IS ILLEGAL')
            return
        else:
            if self.excel_block_list_text.text() != "":
                json_dict['BLOCK'] = self.excel_block_list_text.text().split(',')
            else:
                json_dict['BLOCK'] = []
            if self.excel_matid_list_text.text() != "":
                json_dict['MAT'] = self.excel_matid_list_text.text().split(',')
            else:
                json_dict['MAT'] = []

        if self.excel_top_time_text.text() != "" and self.excel_bottom_time_text.text() != "":
            json_dict['TIME'] = [self.excel_bottom_time_text.text(), self.excel_top_time_text.text()]
        else:
            json_dict['TIME'] = []

        print(json_dict)
        self.excel_textBrowser.append('%s: config info: %s\n' % (now_time, json_dict))
        # update config file
        with open('tqs_excel_config.json', 'w') as json_file:
            json_file.write(json.dumps(json_dict))
        self.excel_textBrowser.append('%s: tqs_excel_config.json is updated sucessfully!\n' % now_time)
        return

    def box_change_config(self):
        with open('box_config.json') as json_file:
            json_dict = json.load(json_file)

        if self.box_block_text.text() != "":
            json_dict['BLOCK'] = self.box_block_text.text().split(',')
        else:
            json_dict['BLOCK'] = []
        if self.box_mat_text.text() != "":
            json_dict['MAT'] = self.box_mat_text.text().split(',')
        else:
            json_dict['MAT'] = []

        if self.box_top_time_text.text() != "" and self.box_bottom_time_text.text() != "":
            json_dict['TIME'] = [self.box_bottom_time_text.text(), self.box_top_time_text.text()]
        else:
            json_dict['TIME'] = []

        self.box_textBrowser.append('%s: config info: %s\n' % (now_time, json_dict))
        # update config file
        with open('box_config.json', 'w') as json_file:
            json_file.write(json.dumps(json_dict))
        self.box_textBrowser.append('%s: box_config.json is updated sucessfully!\n' % now_time)
        return

    def bar_change_config(self):
        with open('bar_tm_config.json') as json_file:
            json_dict = json.load(json_file)
        if self.bar_mat_list_text.text() != "":
            json_dict['MAT'] = self.bar_mat_list_text.text().split(',')
        else:
            json_dict['MAT'] = []

        print(json_dict)
        if self.bar_top_time_text.text() != "" and self.bar_bottom_time_text.text() != "":
            json_dict['TIME'] = [self.bar_bottom_time_text.text(), self.bar_top_time_text.text()]
        else:
            json_dict['TIME'] = []

        self.bar_textBrowser.append('%s: config info: %s\n' % (now_time, json_dict))
        # update config file
        with open('bar_tm_config.json', 'w') as json_file:
            json_file.write(json.dumps(json_dict))
        self.bar_textBrowser.append('%s: bar_tm_config.json is updated sucessfully!\n' % now_time)
        return

    def create_excel(self):
        connection = vf.connect2DB()
        # plot_tqs_summary_data_df(connection, spec_id_dict)
        # plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
        try:
            vf.tqs_summary_data_excel(connection, config_path='tqs_excel_config.json')
            connection.close()
            self.excel_textBrowser.append('%s: Excel generated successfully!\n' % now_time)
        finally:
            connection.close()
            self.excel_textBrowser.append('%s: Error occurred when creating excel! Please call 18116398262 for help.\n' % now_time)
            self.message_box('Error occurred when creating excel')
        return

    def create_bar(self):
        connection = vf.connect2DB()
        # plot_tqs_summary_data_df(connection, spec_id_dict)
        # plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
        try:
            vf.plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
            connection.close()
            self.bar_textBrowser.append('%s: Bar generated successfully!\n' % now_time)
        finally:
            connection.close()
            self.bar_textBrowser.append('%s: Error occurred when creating bar! Please call 18116398262 for help.\n' % now_time)
            self.message_box('Error occurred when creating bar')
        return

    def create_box(self):
        connection = vf.connect2DB()
        # plot_tqs_summary_data_df(connection, spec_id_dict)
        # plot_bar(connection, config_path='bar_tm_config.json', map_path="mapping.xlsx")
        try:
            vf.plot_box(connection, config_path='box_config.json')
            connection.close()
            self.box_textBrowser.append('%s: Boxplot generated successfully!\n' % now_time)
        finally:
            connection.close()
            self.box_textBrowser.append('%s: Error occurred when creating boxplot! Please call 18116398262 for help.\n' % now_time)
            self.message_box('Error occurred when creating boxplot')
        return




app = QtWidgets.QApplication(sys.argv)
hello = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(hello)
hello.show()
app.exec_()


