from time import time
import xlrd
import re
import xlwt
import json



#global variables
all_subjects_dict={}

# global_Cat_Dict={} 


# TODO: dyanimally settings up columns to extract data.

#here row,column number starts at index 0
 #n-1 row is taken in cell value

 #location of file defined below




#all resuable function:



def openSheet(inputsheet):
    loc = (inputsheet)

    wb = xlrd.open_workbook(loc)

    sheet = wb.sheet_by_index(0)



    #to get number of rows in the sheet
    num_rows = sheet.nrows 

    return sheet,num_rows

def buildFinalDic(workperformed_row_index):
    final_Dict={}
    loc = (inputsheet)

    wb = xlrd.open_workbook(loc)

    sheet = wb.sheet_by_index(0)

    #to get number of rows in the sheet
    num_rows = sheet.nrows 
    for r in range(1,num_rows): 
        Work_Performed_ALL=sheet.cell_value(r, workperformed_row_index)
    
        Work_Performed_ALL=re.sub("\d\) ","",Work_Performed_ALL)
        Work_Performed_ALL=re.sub("\d","",Work_Performed_ALL)
        different_categories = Work_Performed_ALL.split(", ")
        # print(different_categories)
        # Creating dictionaries
        for i in different_categories:
            if i not in final_Dict: 
                final_Dict[i]=0.0

    # print(final_Dict)
    return final_Dict

def PrintToExcel(dict_input,workperformed_row_index):
    
    final_Dict=buildFinalDic(workperformed_row_index)
    Reset_dict=final_Dict.copy()

     # ##Creating output sheet
    newsheet=xlwt.Workbook()
    sheet1 = newsheet.add_sheet("TA work Performed")
    # cols=['Subject','Hours']
    

    row = sheet1.row(0)
    row.write(0,'Subject')
    worktype_index_col=1
    for worktype in final_Dict:

        row.write(worktype_index_col,worktype)
        worktype_index_col+=1

         
        

    
    row_number=1
    worktype_index_col=1
    for key in dict_input:
        
        subject=key
        row = sheet1.row(row_number)
        row_number+=1
        #wrting subject name
        row.write(0,subject)
        # print(f'{key}')

        #rest dict foe each subject
        final_Dict=Reset_dict.copy()
        for key2 in dict_input[subject]:
            
           
            work_performed=key2
            Time=dict_input[key][key2]
            
            final_Dict[work_performed]=Time
            # print(f'{key2} {dict_input[key][key2]}')
            
            
            
        
        # print( f'\n {subject}: \n {final_Dict} \n')
         
        for key in final_Dict:
            row.write(worktype_index_col,final_Dict[key])
            worktype_index_col+=1
        worktype_index_col=1
        
        
   
    newsheet.save("Categoryforeachsubject.xls") 

def DataCheckup(sheet,HoursofWork_row_index):
# #place hours work cloumn un cellvalue(m,n ) in n places and check

# #course number column


# #work category column starting from row 2,9 (10th in excel)

    Work_Hours_each_cat=sheet.cell_value(2, HoursofWork_row_index)

    Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)
    Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)

    Work_Hours_each_cat=Work_Hours_each_cat.split(", ")
    print(f'check whether taken input is of right column \n"{Work_Hours_each_cat}\n ====== \n')

##---------------Data checkup ends------------------

def buildAllSubjectsDict(sheet,num_rows):
        
    for r in range(1,num_rows): 
        Coursename=sheet.cell_value(r,4)
        if Coursename not in all_subjects_dict:
            all_subjects_dict[Coursename]={}

# print(all_subjects_dict)

def CategorySeperator(sheet,num_rows,workperformed_row_index,HoursofWork_row_index,totalApporvedHours_row_index):

    for subject in all_subjects_dict:
        global_Cat_Dict={}
        for r in range(1,num_rows): 

            if subject in sheet.cell_value(r,4):
                # print(f'{subject} is present at {r}')




            #building global dict for each category for each particular subject

                

                Work_Performed_ALL=sheet.cell_value(r, workperformed_row_index)
            
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


                ## :: prints row number extracts the workperformed category for each row
                ## :: create a unquie hash map for each row
            
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
                
                #---  /* Debug statements:
                # print('\n printing local list of each category')
                # print(local_list)

                # */---------

                #implement in new python file to improve code structure  
                # ##Creating output sheet
                # newsheet=xlwt.Workbook()
                # sheet1 = newsheet.add_sheet("TA each task")
                # cols=['WorkPerformed','Hours']
                # txt = "Row %s, Col %s"

                # row = sheet1.row(0)
                # row.write(0,'WorkPerformed')
                # row.write(1,'Hours')   

                
                    
                workhour_list=[]
                
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
                        
                    
                    
                    

                    
                # print(f'{subject}  {local_dict_combined}')

                #adding to total:
                for i in local_dict_combined:
                    global_Cat_Dict[i]=global_Cat_Dict[i]+local_dict_combined[i]    
                # print(global_Cat_Dict)        
            

            all_subjects_dict[subject] =global_Cat_Dict       





#inputs:

inputsheet="./FA22TA.xls"
# inputsheet="./backup-old/ACTUALSHEET-old.xls"

#all indexs start from zero
workperformed_row_index=7
#looks like start from zero index
#1) Attendance - TAKING, 2) Attendance - TAKING, 3) Answering e-mails

HoursofWork_row_index=9
# 1) 1.00, 2) 1.00, 3) 0.50

totalApporvedHours_row_index=10
# 3 -- represents the approved usally lesser number


sheet,num_rows=openSheet(inputsheet)
# DataCheckup(sheet,HoursofWork_row_index)
buildAllSubjectsDict(sheet,num_rows)
CategorySeperator(sheet,num_rows,workperformed_row_index,HoursofWork_row_index,totalApporvedHours_row_index)
PrintToExcel(all_subjects_dict,workperformed_row_index)


