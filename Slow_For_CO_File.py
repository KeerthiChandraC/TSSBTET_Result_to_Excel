import os

import pandas as pd

from re import sub
from time import sleep

SLOW_LEARNERS_PERCENT  = 40
FAST_LEARNERS_PERCENT  = 80

INTERNAL_MAX_MARKS = 20
MID1_MAX_MARKS = 20
MID2_MAX_MARKS = 20
EXTERNAL_MAX_MARKS = 40

def set_row_height(row, height):
    #print(row.height)
    row.height = height
    for cell in row.cells:
        cell.height = height

def set_column_width(column, width):
    #print(column.width)
    column.width = width
    for cell in column.cells:
        cell.width = width

    
thisdir = os.getcwd()
print(thisdir)
 

def coderintro():
    screen_time =3
    for i in range(0,screen_time):
                os.system('cls')
                print()
                print("********************************************************************************")
                print()
                print("developed by Keerthi Chandra C")
                print("Lecturer in ECE , GMRPW Karimnagar")
                print("for constructive criticism and suggestions reach me at keerthichand.c@gmail.com")
                print('''
disclaimer:
    *These are not actual Results
    *Use this at your own discretion
    *for actual TimeTables please visit TS-SBTET Website
    *OUTPUTS from this program are not related to TS-SBTET
    *This program is developed only to help ease work but not to misguide or harm in 
        any possible ways and with no ill intent
    ''')
                print("******************************************************************************")
                if screen_time-i-1>0:
                        print(f'starting in {screen_time-i-2}s .......')
                sleep(1)
    input("press ENTER to start.....")
    os.system('cls')



def readxl(file_name,outdir):
        print(f"Reading {file_name}.......",end='')
        df = pd.read_excel(file_name, header=[1], sheet_name="Sheet")
        
        cols  = list(df.columns.values)
        subList = [x for x in cols if "_" in x and "-" in x  ]
        subList = [x.split("_")[0] for x in  subList]
        subList = [*set(subList)]
        subList.sort()
        for sub in subList:
            writer = pd.ExcelWriter(f'{outdir}/{sub}.xlsx')
            col = [ x for x in cols if sub in x] 
            #print(col)
            df_over = df[["PIN","NAME"]+col]
            df_over.insert(loc=0, column="Sl.No.", value=df_over.reset_index().index+1)
            df_over.to_excel(writer,f'{sub}_overview',index=False)


            

            
            for c in col:
                
                if "_status" in c:

                    rslt_df = df
                    pas = rslt_df[f'{c}'].value_counts()['PASS']
                    total = rslt_df[f'{c}'].count()
                    fail =  total - pas
                    data_dict = {f"{sub}":["Total Students:","No of Passed Students:","No of Failed Students:", "Pass Percentage:"],
                            f"Result Analysis":[f"{total}",f"{pas}",f"{fail}", f"{round((pas/total)*100,2)}"]}
                    
                    data  = pd.DataFrame(data_dict)
                    data.to_excel(writer,f'{sub}_Result Analysis',index=False)
            for c in col:
                if "_mid1" in c:
                    
                    rslt_df = df.loc[(df[c] < (0.01* SLOW_LEARNERS_PERCENT * MID1_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'MID1_DULL',index=False)

                    rslt_df = df.loc[(df[c] >= (0.01* FAST_LEARNERS_PERCENT * MID1_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'MID1_TOP',index=False)

                if "_mid2" in c:
                    
                    rslt_df = df.loc[(df[c] < (0.01* SLOW_LEARNERS_PERCENT * MID2_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'MID2_DULL',index=False)

                    rslt_df = df.loc[(df[c] >= (0.01* FAST_LEARNERS_PERCENT * MID2_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'MID2_TOP',index=False)

                if "_intr" in c:
                    
                    rslt_df = df.loc[(df[c] < (0.01* SLOW_LEARNERS_PERCENT * INTERNAL_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'intr_DULL',index=False)

                    rslt_df = df.loc[(df[c] >= (0.01* FAST_LEARNERS_PERCENT * INTERNAL_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'intr_TOP',index=False)
                if "_ext" in c:
                    
                    rslt_df = df.loc[(df[c] < (0.01* SLOW_LEARNERS_PERCENT * EXTERNAL_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'ext_DULL',index=False)

                    rslt_df = df.loc[(df[c] >= (0.01* SLOW_LEARNERS_PERCENT * EXTERNAL_MAX_MARKS ))]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'ext_TOP',index=False)
                if "_total" in c:
                    
                    rslt_df = df.loc[(df[c] < 14)]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index+1)
                    rslt_df.to_excel(writer,f'total_DULL',index=False)

                    rslt_df = df.loc[(df[c] >= 35)]
                    rslt_df = rslt_df[["PIN","NAME",f'{c}']]
                    rslt_df.insert(loc=0, column="Sl.No.", value=rslt_df.reset_index().index + 1)
                    rslt_df.to_excel(writer,f'total_TOP',index=False)


            writer.close()

        
            
           
if __name__=='__main__':
    count = 0
    coderintro()
    #instructions_info()
    dir_list = os.listdir(f'{thisdir}/INPUT')
    for file in dir_list:
        if (file.endswith(".xlsx") or file.endswith(".xlx")) :
            print(f'{thisdir}/INPUT/{file}')
            print(f"Converting {file}")
            readxl(f'{thisdir}/INPUT/{file}',f'{thisdir}/OUTPUT')
            print(f"Conversion copleted......")
            count +=1
    
    if count >0:
        print("All files are saved in Outputs folder")
        input("Press Enter to Exit")
    else:
        print("No excel files found")
        input("Press Enter to Exit")
