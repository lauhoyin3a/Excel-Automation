import time
import pyautogui
import pandas as pd
import os
import pyperclip

vba_file_path=r'C:\Users\Tom\Downloads\VBA script.txt'
with open(vba_file_path,'r') as file:
    vba_script=file.read()
print(vba_script)

def pd_checking(excel_file_path,field_name, field_type):
    df=pd.read_excel(excel_file_path)
    match_field_name_len=len(df[df['name']==field_name])
    try:
        match_field_type_len=len(df[field_type])
    except:
        print("key error")
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
    print(vba_script)
    pyautogui.typewrite(vba_script,interval=0.0007)
    time.sleep(3)
def run_vba_script(field_name,column_name,value):
    time.sleep(3)
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
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('ctrl','s')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    print(f'modified {field_name}-{column_name} to {value}')

field_name=input('target field name to be modify: ')
column_name=input('target column name to be modify: ')
value=input('input value:')

process=input(f'Ready to modify {column_name} field from {field_name} in excels, enter [y/n]: ')
if process == 'n':
    raise SystemExit


excel_file_path_list=[]
excel_file_path_list.append(r'C:\Users\Tom\abc.xlsx')


for excel_file_path in excel_file_path_list:
    data_exist = pd_checking(excel_file_path, field_name, column_name)
    if not data_exist:
        print(f"{field_name} doesnot exist")
        continue
    paste_vba_script(vba_script,excel_file_path)
    run_vba_script(field_name,column_name,value)
