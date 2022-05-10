from datetime import timedelta, date
import openpyxl as xl
import os
import glob
import pandas as pd
import numpy as np

#======================
# Importing Control File with Directories

path_source = "C:\\Users\\" + os.environ["USERNAME"] + "\\"

dir_control = glob.glob(path_source +  "add dir pathway here")

df_dir = pd.read_excel(dir_control[0], sheet_name = "Control")

cc_list = df_dir["ClientCode"].to_numpy()


#======================

print("==========================================")
print(cc_list) # Prints a list of clients available to be aggregated
print("==========================================")

print("***********")
print("For Trade Files, look at the above list and input the client code with appropriate T placement \n")


client_code = str(input("Please Enter Fund Code from Above list of names: \n")) # Request input from user on which client to analyse.

#======================
# This section checks whether the string entered into the console matches the code within the list to prevent a directory error being produced

cc_check = False

while cc_check is False:  

    if client_code in cc_list: # Attempts to match entered value with values in list
        cc_check = True # Change to True to exit while loop
        
    else:
        print("***********")
        print("Code entered is NOT present in list: \n")
        client_code = str(input("Re-enter client code: ")) # If the client code does not match, requests user to re-input file.
        print("***********")
  
#======================

print("***********")
# Requests input from user for a date range in which the aggregation takes place. 
# Need to create error catching for Y, M, D.

crnt_year = date.today().year
d_lst = [i for i in range(1,32)]

# could add a section where month is usedto find range of dates in that month

m_lst = [j for j in range(1,13)]
y_lst = [k for k in range(2018,crnt_year + 1)]


def mdycheck(lst, date):
    complete = False
    
    while complete is False:
        
        print(lst) # Print list of inputs to user
        start = input("Enter a value for the "+ date +" : \n") # Ask user to input 
        int_check = start.isnumeric() # check to see if an integer
        print("************************************** \n") 
        
        if int_check is True and int(start) in lst: # breaks loop if both values are true
            break
        
        
        if int_check is True and int(start) not in lst: 
            print("Value entered is not in the list: \n")
            print("**************************************")
            continue
        
        if int_check is False:
            print("Valued entered is not an Integer!")
            print("**************************************")
            continue
        

    return(int(start))   


print("***********")
print("START Date \n")
yyyy_s = mdycheck(y_lst,"Year" )
mm_s = mdycheck(m_lst ,"Month")
dd_s = mdycheck(d_lst,"Day")
print("***********")
print("***********")
print("Enter END Date Date \n")
yyyy_e = mdycheck(y_lst,"Year" )
mm_e = mdycheck(m_lst ,"Month")
dd_e = mdycheck(d_lst,"Day")
print("***********")


#======================
# Adding the path source and date format to variables by using the code entered by the user, and cross-referencing against the directory spreadsheet

file_path = df_dir.loc[df_dir["ClientCode"] == client_code, "SourcePath"].to_list()[0]
date_format = df_dir.loc[df_dir["ClientCode"] == client_code, "DateFormat"].to_list()[0]


start_dt = date(yyyy_s,mm_s,dd_s) # Edit this input to change range of dates that the code will find and aggregate
end_dt = date(yyyy_e,mm_e,dd_e)
weekdays = [5,6] 
dates = []

# This funcion finds the working days within the date range, and appends them to a list in yyyymmdd format. 

def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)


for dt in daterange(start_dt, end_dt): # Calling function
    if dt.weekday() not in weekdays: # Removes weekends from list, to only have business days.
        d = dt.strftime(str(date_format))
        dates += [d]

#======================
# Setting up dataframes and addition paths for use

t_dir = path_source + "\dir redacted\" + file_path
df = []
df_master = []

#======================

for i in range(len(dates)): # Iterating through the dates to create directory to search for the file.
    
    temp_dir = t_dir % dates[i] # Adding date to temporary directory
    
    temp_per = str(np.round(i/int(len(dates)) * 100)) # Creating percentage for use in progress bar
    
    
    print("********************")   
    print("Code is Running \n")
    print("Progress: ", temp_per + "%")
    print("********************")
    
    try: # Attempting to open directory created above
        
        f = open(temp_dir)
    
    except PermissionError: # Checks to see whether the file is open
        print("**************************** \n")
        print("The following file is open: \n")
        print(temp_dir, "\n")
        print("Please close this file and re-run the script")
        print("****************************")
        print("CODE HAS ENDED")
        print("****************************")
        break
    
    except FileNotFoundError: # Checks the directory pathway, and whether it can be found
        
        print("****************************")
        print("Cannot Open File: \n")
        print(temp_dir, "\n")
        print("This will NOT be added too the dataframe \n")
        print("****************************")
        continue

    if temp_dir[-4:] == ".csv": # File types vary - checking for .csv or .xlsx filetype for appropriate calling of a function
        df.append(pd.read_csv(temp_dir)) # Appending data to temporary dataframe
        df_master = pd.concat(df, ignore_index=True) # Concatinating to master dataframe
        
    elif temp_dir[-4:] in ['.xls', 'xlsx', 'xlsm']:
        
        df.append(pd.read_excel(temp_dir))
        df_master = pd.concat(df, ignore_index=True)
    
    else: # Prints directory that cannot be accessed, but carries on with the iteration in case of known users.
        
        print("****************************")
        print("File Type Not Accessible: \n")
        print(temp_dir, "\n")
        print("This will NOT be added to the dataframe \n")
        print("****************************")
        continue

    
# Creating a excel spreadsheet of aggregated data and naming is after the client code, into a known file directory
df_master.to_excel (r"C:\\Users\\" + os.environ['USERNAME'] + "add dir pathway here" % client_code, index = None, header=True)


print("***************************")
print("CODE HAS COMPLETED \n")
print("Check the Aggregation Folder for the spreadsheet \n")
print("***************************")     



#======================   

