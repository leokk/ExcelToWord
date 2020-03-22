# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '1.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(335, 176)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.select_button = QtWidgets.QPushButton(self.centralwidget)
        self.select_button.setGeometry(QtCore.QRect(220, 30, 75, 23))
        self.select_button.setObjectName("select_button")
        self.result_label = QtWidgets.QLabel(self.centralwidget)
        self.result_label.setGeometry(QtCore.QRect(120, 120, 47, 13))
        self.result_label.setObjectName("result_label")
        self.header_label = QtWidgets.QLabel(self.centralwidget)
        self.header_label.setGeometry(QtCore.QRect(20, 60, 47, 13))
        self.header_label.setObjectName("header_label")
        self.calculate_button = QtWidgets.QPushButton(self.centralwidget)
        self.calculate_button.setGeometry(QtCore.QRect(220, 80, 75, 23))
        self.calculate_button.setObjectName("calculate_button")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.select_button.setText(_translate("MainWindow", "Select"))
        self.result_label.setText(_translate("MainWindow", "TextLabel"))
        self.header_label.setText(_translate("MainWindow", "TextLabel"))
        self.calculate_button.setText(_translate("MainWindow", "PushButton"))

