# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'COPITTATimeSheet.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QFileDialog
from TAWorkedPerformed import *




class Ui_MainWindow(object):


    def CompileSheets():
        

        # TAWork Performed
        
        TAWork
        # sheet,num_rows=openSheet(inputsheet)

        # dataCheck(sheet)
        # BuildingDict(num_rows)
        # CategorySperator(num_rows,workperformed_row_index,HoursofWork_row_index,totalApporvedHours_row_index)
        # print(global_Cat_Dict)        
        # CreatingOutputSheet()



    def upload(self):
        fullPath = QFileDialog.getOpenFileName(self.centralwidget,"Open File", "../", "All Files (*)")
        # fname = QFileDialog.getOpenFileName(self, "Open File", "./test_images/", "All Files (*)")
        # object name is stored in pixmap1
        # self.pixmap1=QPixmap(fullPath[0]) 
        self.filepath=fullPath[0]
        # open image
        print(self.filepath)

        self.CompileSheets()
     
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1005, 713)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(130, 50, 113, 21))
        self.lineEdit.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 10, 101, 16))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(30, 50, 91, 16))
        self.label_2.setObjectName("label_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(390, 50, 113, 21))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(270, 50, 121, 20))
        self.label_3.setObjectName("label_3")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setGeometry(QtCore.QRect(800, 50, 113, 21))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(630, 50, 131, 16))
        self.label_4.setObjectName("label_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_5.setGeometry(QtCore.QRect(390, 130, 113, 21))
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(270, 130, 111, 20))
        self.label_5.setObjectName("label_5")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_6.setGeometry(QtCore.QRect(140, 130, 113, 21))
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(10, 130, 111, 16))
        self.label_6.setObjectName("label_6")
        self.lineEdit_8 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_8.setGeometry(QtCore.QRect(800, 120, 113, 21))
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(640, 120, 121, 16))
        self.label_8.setObjectName("label_8")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget,clicked=lambda: self.upload())
        self.pushButton.setGeometry(QtCore.QRect(140, 290, 161, 32))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(430, 300, 161, 32))
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(130, 190, 113, 21))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(20, 190, 91, 16))
        self.label_7.setObjectName("label_7")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "COPIT TA Time Sheet Tool"))
        self.lineEdit.setText(_translate("MainWindow", "5"))
        self.label.setText(_translate("MainWindow", "Column Set up"))
        self.label_2.setText(_translate("MainWindow", "Course Name"))
        self.lineEdit_3.setText(_translate("MainWindow", "6"))
        self.label_3.setText(_translate("MainWindow", "Coordinator Name"))
        self.lineEdit_4.setText(_translate("MainWindow", "8"))
        self.label_4.setText(_translate("MainWindow", "Work Performed"))
        self.lineEdit_5.setText(_translate("MainWindow", "10"))
        self.label_5.setText(_translate("MainWindow", "Hours of Work"))
        self.lineEdit_6.setText(_translate("MainWindow", "9"))
        self.label_6.setText(_translate("MainWindow", "Work Description"))
        self.lineEdit_8.setText(_translate("MainWindow", "11"))
        self.label_8.setText(_translate("MainWindow", "Total Course Hours"))
        self.pushButton.setText(_translate("MainWindow", "Upload Excel Sheet"))
        self.pushButton_2.setText(_translate("MainWindow", "Submit"))
        self.lineEdit_2.setText(_translate("MainWindow", "1"))
        self.label_7.setText(_translate("MainWindow", "Name"))





if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())