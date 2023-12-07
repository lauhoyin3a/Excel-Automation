import time
import pyautogui
import pandas as pd
import os
import pyperclip
import json
import shutil

vba_file_path=r'C:\Users\Tom\Downloads\VBA script.txt'
with open(vba_file_path,'r') as file:
    vba_script=file.read()
checking_path=r"C:\Users\Tom\Downloads\checking.txt"

with open(checking_path,'r') as file:
    checking_text=file.read()

excel_file_path_list=[]


def copy_file(input_path,output_path):
    shutil.copy(input_path,output_path)

def pd_checking(excel_file_path,field_name, field_type):
    df=pd.read_excel(excel_file_path)
    match_field_name_len=len(df[df['name']==field_name])
    try:
        match_field_type_len=len(df[field_type])
    except:

        print(f"key error for field name: {field_name}")
        return False

    if match_field_name_len>0:
        return True
    return False


def paste_vba_script(vba_script,excel_file_path):

    os.startfile(excel_file_path)
    time.sleep(2)

    pyautogui.hotkey('alt', 'f11')
    time.sleep(1)
    pyautogui.hotkey('alt', 'i')
    time.sleep(1)
    pyautogui.hotkey('m')
    time.sleep(1)
    pyautogui.moveTo(579,364)
    pyautogui.click()
    pyautogui.typewrite(vba_script,interval=0.0007)
    time.sleep(2)

def run_vba_script(field_name,column_name,value):
    time.sleep(1)
    pyautogui.hotkey('f5')
    time.sleep(2)
    pyautogui.typewrite(field_name)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.typewrite(column_name)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.typewrite(value)
    pyautogui.press('enter')

    pyautogui.press('enter')

    pyautogui.press('enter')
    time.sleep(1)
    # pyautogui.hotkey('ctrl','s')
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(2)
    print(f'modified {field_name}-{column_name} to {value}')
def exit_excel():
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)
    pyautogui.press('enter')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')



    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
def read_csv_input(file_path):
    data = pd.read_excel(file_path)
    print(data.head())
    name_list,field_list,value_list=[],[],[]
    print(data['name'])
    for index, row in data.iterrows():
        name_list.append(row['name'])
        field_list.append(row['target field'])
        value_list.append(row['value'])
    print(name_list)
    print(field_list)
    print(value_list)
    return (name_list,field_list,value_list)

def read_file_path(file_path):
    f = open(file_path)
    data = json.load(f)
    for path in data['excel_file_path_list']:
        excel_file_path_list.append(path)




name_list,field_list,value_list=read_csv_input(r"C:\Users\Tom\Downloads\value.xlsx")
#field_name=input('target field name to be modify: ')
#column_name=input('target column name to be modify: ')
#value=input('input value:')
print('Please set keyboard input layout as US')
for i in range(len(name_list)):
    print(f'Ready to modify {name_list[i]} field from {field_list[i]} in excels into values: {value_list[i]} ')

process=input('Confirm to modify the above changes, enter [y/n]: ')
if process == 'n':
    raise SystemExit


read_file_path('staging_survey.json')
#read_file_path("test_survey.json")
#for excel_file_path in excel_file_path_list:
#      copy_file(excel_file_path,r"C:\Users\Tom\Downloads\survey_backup")
# exit()
for excel_file_path in excel_file_path_list:
    paste_vba_script(vba_script, excel_file_path)
    for i in range(len(name_list)):

        data_exist = pd_checking(excel_file_path, name_list[i], field_list[i])
        if not data_exist:
            print(f"{name_list[i]} doesnot exist")
            continue

        run_vba_script(name_list[i],field_list[i],value_list[i])
    exit_excel()
