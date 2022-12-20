import xlrd
import re
import xlwt
import json
global_Cat_Dict={} 


# TODO: dyanimally settings up columns to extract data.

#here row,column number starts at index 0
 #n-1 row is taken in cell value

 #location of file defined below


class TAworksheet(self):

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


        #all indexs start from zero
        workperformed_row_index=7
        #looks like start from zero index
        #1) Attendance - TAKING, 2) Attendance - TAKING, 3) Answering e-mails

        HoursofWork_row_index=9
        # 1) 1.00, 2) 1.00, 3) 0.50

        totalApporvedHours_row_index=10
        # 3 -- represents the approved usally lesser number


      

        def setupInputSheet(self,filepath):
            self.inputsheet=filepath


# sheet,num_rows=openSheet(inputsheet)

# dataCheck(sheet)
# BuildingDict(num_rows)
# CategorySperator(num_rows,workperformed_row_index,HoursofWork_row_index,totalApporvedHours_row_index)
# print(global_Cat_Dict)        
# CreatingOutputSheet()
