import xlrd
import re
import xlwt
import json
global_Cat_Dict={} 
 
loc = ("./Spring2022 TA Time-edited.xls")

wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_index(0)

#to get number of rows in the sheet
num_rows = sheet.nrows - 1

Work_Hours_each_cat=sheet.cell_value(2, 9)
Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)
Work_Hours_each_cat=re.sub("\d\) ","",Work_Hours_each_cat)

Work_Hours_each_cat=Work_Hours_each_cat.split(", ")
# print(Work_Hours_each_cat)



#building global dict for each category
for r in range(1,num_rows): 
    Work_Performed_ALL=sheet.cell_value(r, 8)
   
    Work_Performed_ALL=re.sub("\d\) ","",Work_Performed_ALL)
    Work_Performed_ALL=re.sub("\d","",Work_Performed_ALL)
    different_categories = Work_Performed_ALL.split(", ")
    # print(different_categories)
    # Creating dictionaries
    for i in different_categories:
        if i not in global_Cat_Dict: 
            global_Cat_Dict[i]=0.0

print(global_Cat_Dict)
#total activity count
print(len(global_Cat_Dict))


for r in range(1,num_rows): 
    local_list=[]
    print(f'rownumber={r}')
    Work_Performed_ALL=sheet.cell_value(r, 8)
    Work_Performed_ALL=re.sub("\d\) ","",Work_Performed_ALL)
    Work_Performed_ALL=re.sub("\d","",Work_Performed_ALL)
    different_categories = Work_Performed_ALL.split(", ")
    
    for i in different_categories:
        # if i not in local_dict: 
            local_list.append(i)   

    print(local_list)        
    workhour_list=[]
    Work_Hours_each_cat=sheet.cell_value(r, 9)
    Work_Hours_each_cat=re.sub("\d*\) ","",Work_Hours_each_cat)
   
    Work_Hours_each_cat=Work_Hours_each_cat.split(", ")
    
    for i in Work_Hours_each_cat:
            workhour_list.append(i) 

    #prints hours for each row of work category
    print(workhour_list)        
    local_dict_combined={}
    
    for i in different_categories:
        if i not in local_dict_combined: 
            local_dict_combined[i]=0.0   
    
    check_Total_Course_hours=sheet.cell_value(r, 13)
    
    r_local_in_cell=0
    prev=[]




    #debug:
    LV_check_total=0.0

    # prev.append(Work_Hours_each_cat[0])
    for i in local_list:

        if(i in prev):
            #activity already seen add up
            local_dict_combined[i]=float(local_dict_combined[i])+float(workhour_list[r_local_in_cell])
            
            #debug:
            LV_check_total=LV_check_total+float(workhour_list[r_local_in_cell])


        else:
            #activity not see add entry
            local_dict_combined[i]=float(workhour_list[r_local_in_cell])
            prev.append(i)

            #debug:
            LV_check_total=LV_check_total+float(workhour_list[r_local_in_cell])

            


        r_local_in_cell=r_local_in_cell+1

        #debug:
        # check for the total hours and local value
        # 

    if(float(check_Total_Course_hours)!=LV_check_total):
        print(f'{check_Total_Course_hours} not equal to {LV_check_total}')
            
        
        
        

        
    print(local_dict_combined)

    #adding to total:
    for i in local_dict_combined:
         global_Cat_Dict[i]=global_Cat_Dict[i]+local_dict_combined[i]    
    print(global_Cat_Dict)        
print(global_Cat_Dict)        


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

# # T=global_Cat_Dict.items()
# print(T)
# print(T[0])
# print(T[2])
# print(T[3])
# print(T[10])