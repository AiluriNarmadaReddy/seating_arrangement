import pandas as pd
import time
import os.path
import openpyxl
import sys
bname = {
        '01': "CIV",
        '02': "EEE",
        '03': "MECH",
        '04': "ECE",
        '05': "CSE",
        '12': "IT",
        '21': "AERO",
        '66': "CSM",
        '27': "XYZ"
    }

def get_branch_name(ht_no):
    if len(ht_no) == 10:
        branch_code = ht_no[6:8]
        if branch_code in bname:
            return bname[branch_code]
    return ""


excel_file_name = sys.argv[1]
#excel_file_name=input("enter the file name")
year = time.strftime("%y", time.localtime())
mon = time.strftime("%b", time.localtime())
date = time.strftime("%d", time.localtime())

df_list = []
for sheet_name in pd.read_excel(excel_file_name, sheet_name=None):
    df_list.append(pd.read_excel(excel_file_name, sheet_name=sheet_name))
student_list = pd.concat(df_list)

student_list['Branch Name'] = student_list['Hallticketno'].apply(get_branch_name)
student_list['index'] = range(1, len(student_list)+1)
no_of_students=len(student_list)
branch_dict = {branch: [] for branch in student_list['Branch Name'].unique()}
for index, row in student_list.iterrows():
    branch = row['Branch Name']
    student_id = row['Hallticketno']
    if branch in branch_dict:
        branch_dict[branch].append(student_id)
branch_wise_list=list(branch_dict.values())
for i, lst in enumerate(branch_wise_list):
    if i == 0:
        even_list = lst.copy()
        continue
    if i == 1:
        odd_list = lst.copy()
        continue
    if len(even_list) > len(odd_list):
        odd_list += lst
    else:
        even_list += lst

if len(even_list)<len(odd_list):
    for i in range (len(even_list),len(odd_list)):
        even_list.append(None)
elif len(odd_list)<len(even_list):
    for i in range(len(odd_list),len(even_list)):
        odd_list.append(None)
#print(len(even_list),len(odd_list))
rooms = pd.read_excel('rooms.xlsx')
no_of_seats=rooms.apply(lambda x: x['Rows'] * x['Columns'], axis=1).sum()
if (len(even_list) + len(odd_list)) - no_of_seats <= 0:
         print("no more rooms required")
else:
         print("still need more rooms")
if not os.path.exists(date+"_"+mon+"_"+year+"_"+'all_roll.xlsx'):
    with pd.ExcelWriter(date+"_"+mon+"_"+year+"_"+'all_roll.xlsx') as writer:
        student_list.to_excel(writer, sheet_name='Sheet1', index=False)
else:
    openpyxl.load_workbook(date+"_"+mon+"_"+year+"_"+'all_roll.xlsx')

