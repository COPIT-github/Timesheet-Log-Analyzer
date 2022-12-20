# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'COPITTATimeSheet.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QFileDialog
import xlrd
import re
import xlwt
import json

global_Cat_Dict={} 




class TAWorkPerformed():
    
    def openSheet(self,inputsheet):
        
        #inputsheet
        loc = (inputsheet)
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)

        #to get number of rows in the sheet
        num_rows = sheet.nrows 
        return sheet,num_rows

    def dataCheck(self,sheet):
            
        #place hours work cloumn un cellvalue(m,n ) in n places and check

        Work_Hours_each_cat=sheet.cell_value(2, 9)

        Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)
        Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)

        Work_Hours_each_cat=Work_Hours_each_cat.split(", ")
        print(f'check whether taken input is of right column \n"{Work_Hours_each_cat}\n ====== \n')






    def CreatingOutputSheet():
        ##Creating output sheet
        newsheet=xlwt.Workbook()
        sheet1 = newsheet.add_sheet("TA work Performed Final")
        cols=['WorkPerformed','Hours']
        txt = "Row %s, Col %s"

        row = sheet1.row(0)
        row.write(0,'WorkPerformed')
        row.write(1,'Hours')

        index=1
        work_per=0
        Hours_per=1
        for key,value in global_Cat_Dict.items():
            
            row = sheet1.row(index)

            row.write(work_per,key)
            row.write(Hours_per,value)
            index=index+1

        # T=json.dumps(global_Cat_Dict)
        newsheet.save("TA work Performed.xls")
        

    def BuildingDict(num_rows):
    #building global dict for each category
        for r in range(1,num_rows): 
            Work_Performed_ALL=sheet.cell_value(r, 7)
        
            Work_Performed_ALL=re.sub("\d\) ","",Work_Performed_ALL)
            Work_Performed_ALL=re.sub("\d","",Work_Performed_ALL)
            different_categories = Work_Performed_ALL.split(", ")
            # print(different_categories)
            # Creating dictionaries
            for i in different_categories:
                if i not in global_Cat_Dict: 
                    global_Cat_Dict[i]=0.0




    # print(global_Cat_Dict)
    #total activity count
    # print(len(global_Cat_Dict))


    def CategorySperator(num_rows,workperformed_row_index,HoursofWork_row_index,totalApporvedHours_row_index):
    #Builds list for each category using a List
    ## :: prints row number extracts the workperformed category for each row
    ## :: create a unquie hash map for each row
        for r in range(1,num_rows): 
            local_list=[]
            #print row number
            # print(f'rownumber={r}')
            Work_Performed_ALL=sheet.cell_value(r, workperformed_row_index)
            Work_Performed_ALL=re.sub("\d\) ","",Work_Performed_ALL)
            Work_Performed_ALL=re.sub("\d","",Work_Performed_ALL)
            different_categories = Work_Performed_ALL.split(", ")
            
            for i in different_categories:
            #add all without removing duplicates 
                    local_list.append(i)   
            


        

            
                
            workhour_list=[]
            #row looks like 1) 1.00, 2) 1.00, 3) 0.50
            #removes all the indexs from the work row but those are they key for us to add up the data attendance taking, attendance taking
            Work_Hours_each_cat=sheet.cell_value(r, HoursofWork_row_index)
            Work_Hours_each_cat=re.sub("\d*\) ","",Work_Hours_each_cat)
        
            Work_Hours_each_cat=Work_Hours_each_cat.split(", ")
            
            for i in Work_Hours_each_cat:
                    workhour_list.append(i) 

            #Debug statement:
            # prints hours for each row for each work category
            #Debug : 
            # print(workhour_list)
            
                    
            local_dict_combined={}
            
            #considers all the extracted column data in different categories
            for i in different_categories:
                if i not in local_dict_combined: 
                    #initializing each category for each row under zero
                    local_dict_combined[i]=0.0   

            
            check_Total_Course_hours=sheet.cell_value(r, totalApporvedHours_row_index)
            # print(check_Total_Course_hours)
            
            r_local_in_cell=0
            #prev for maintaing the lookup while considering the hours of working and work performed
            




            #debug:
            LV_check_total=0.0
            prev=[]
            #we use work performed row to look and utilize the workhour_list to add things
            
            for i in local_list:
                
                
                if(i in prev):
                    #DEBUG:
                    #activity already seen add up
                    # print(f'already seen this up')

                    local_dict_combined[i]=float(local_dict_combined[i])+float(workhour_list[r_local_in_cell])
                    
                    #debug:
                    LV_check_total=LV_check_total+float(workhour_list[r_local_in_cell])


                else:
                    
                    #activity not see add entry
                    #Debug statement
                    # print(f'index={i}')
                    #converting the worklist at each index 


                    local_dict_combined[i]=float(workhour_list[r_local_in_cell])
                    #appending the key
                    prev.append(i)

                    #debug:
                    LV_check_total=LV_check_total+float(workhour_list[r_local_in_cell])
                    # print(LV_check_total)

                # Debug statement
                # print(prev)   

                #adding up the index for workhour_list
                r_local_in_cell=r_local_in_cell+1


                #debug:
                # check for the total hours and local value
                # print(f'here={check_Total_Course_hours}')
                # m=float(check_Total_Course_hours)
                # print(m)
                # print(float(LV_check_total))
                
            if(float(check_Total_Course_hours)!=float(LV_check_total)):
                print(f'{check_Total_Course_hours} not equal to {LV_check_total}')
                    
                
                
                

                
            # print(local_dict_combined)

            #adding to total:
            for i in local_dict_combined:
                global_Cat_Dict[i]=global_Cat_Dict[i]+local_dict_combined[i]    
            # print(global_Cat_Dict)        







    def setupInputSheet(self,filepath):
        self.inputsheet=filepath
        print(self.inputsheet)
            #all indexs start from zero
        workperformed_row_index=7
        #looks like start from zero index
        #1) Attendance - TAKING, 2) Attendance - TAKING, 3) Answering e-mails

        HoursofWork_row_index=9
        # 1) 1.00, 2) 1.00, 3) 0.50

        totalApporvedHours_row_index=10
        # 3 -- represents the approved usally lesser number




class Ui_MainWindow(object):


    

    def CompileSheets(self):
        

        # TAWork Performed
        TWP=TAWorkPerformed()
        TWP.setupInputSheet(self.filepath)
        
        
     

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