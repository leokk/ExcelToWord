import sys

from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QFileDialog
from qtpy import QtCore

from FileMaker import Filemaker


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = "Title"
        self.top = 100
        self.left = 100
        self.width = 640
        self.height = 480
        self.filename = None
        self.InitWindow()

    def InitWindow(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setupUi()
        self.show()

    def setupUi(self):
        self.setObjectName("MainWindow")
        self.setFixedSize(335, 176)

        self.select_button = QPushButton(self)
        self.select_button.setGeometry(QtCore.QRect(220, 30, 75, 23))
        self.select_button.setObjectName("select_button")
        self.select_button.setText("Choose")
        self.select_button.clicked.connect(self.openFileDialog)

        self.result_label = QLabel(self)
        self.result_label.setGeometry(QtCore.QRect(40, 120, 100, 23))
        self.result_label.setObjectName("result_label")
        self.result_label.setText('Waiting for start')

        self.header_label = QLabel(self)
        self.header_label.setGeometry(QtCore.QRect(40, 30, 100, 23))
        self.header_label.setObjectName("header_label")
        self.header_label.setText('Choose.doc file')

        self.calculate_button = QPushButton(self)
        self.calculate_button.setGeometry(QtCore.QRect(220, 120, 75, 23))
        self.calculate_button.setObjectName("calculate_button")
        self.calculate_button.setText("Calculate")
        self.calculate_button.clicked.connect(self.calculate)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.select_button.setText(_translate("MainWindow", "Select"))
        self.result_label.setText(_translate("MainWindow", "TextLabel"))
        self.header_label.setText(_translate("MainWindow", "TextLabel"))
        self.calculate_button.setText(_translate("MainWindow", "PushButton"))

    def calculate(self):
        print(self.filename)
        if self.filename is None:
            self.result_label.setText('Select file first')
        else:
            self.result_label.setText('In progress')
            self.update()
            filemaker = Filemaker(self.filename)

            filemaker.generte()
            self.result_label.setText('Done')
            self.update()

    def openFileDialog(self):
        self.filename = QFileDialog.getOpenFileName(filter="MS - Excel (*.xls; *.xlsx)")[0]


app = QApplication(sys.argv)
window = Window()

sys.exit(app.exec())
