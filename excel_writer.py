import time
import pyautogui
import pandas as pd
import os
import pyperclip

vba_file_path=r'C:\Users\Tom\Downloads\VBA script.txt'
with open(vba_file_path,'r') as file:
    vba_script=file.read()
checking_path=r"C:\Users\Tom\Downloads\checking.txt"

with open(checking_path,'r') as file:
    checking_text=file.read()

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
    print(vba_script)
    pyautogui.typewrite(vba_script,interval=0.0007)
    time.sleep(2)
    # pyautogui.hotkey('ctrl','a')
    # time.sleep(1)
    # pyautogui.hotkey('ctrl','c')
    # time.sleep(1)
    # print(pyperclip.paste())
    # if pyperclip.paste()==pyperclip.copy(checking_path):
    #     print("same")
    #time.sleep(10)
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
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(1)
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
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')


    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
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


excel_file_path_list=[]
excel_file_path_list.append(r'C:\Users\Tom\ArcGIS\My Survey Designs\GOMPS_Event_Misc 122\GOMPS_Event_Misc 122.xlsx')
#excel_file_path_list.append(r'C:\Users\Tom\ArcGIS\My Survey Designs\GOMPS_Typhoon_Shelter 3\GOMPS_Typhoon_Shelter 3.xlsx')

# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\1f898ea71b894bedb43b3face5b23faa\GOMPS_Event_Misc.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\f646d8b9480549638137236afc417364\GOMPS_Typhoon_Shelter.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\f11f024cf8e8410bb1e0934ecc57e39c\GOPMS_Hopsital.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\d15fdf8320ae46c4bd286dcd2d988571\GOMPS_Restricted_Access.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\994e964647f14fa297eeb30e80f922a8\GOMPS_Railway_Development.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\71ec010703a14a8d855dda625f3ddc10\GOMPS_Dangerous_Goods.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\d7655e7434aa4f1ca9459828b1a2d041\GOPMS_Ambulance.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\d6de5b3acd7642668aa3e52307fbfe5a\GOMPS_Hotel.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\7ae799842cdb478da58f67ac43683c76\GOMPS_Tunnel.xlsx")
# excel_file_path_list.append(r"C:\Users\Tom\ArcGIS\My Survey Designs\815511cb28dd480a99ca418c6a02149b\GOMPS_view_edit.xlsx")
for excel_file_path in excel_file_path_list:
    paste_vba_script(vba_script, excel_file_path)
    for i in range(len(name_list)):

        data_exist = pd_checking(excel_file_path, name_list[i], field_list[i])
        if not data_exist:
            print(f"{name_list[i]} doesnot exist")
            continue

        run_vba_script(name_list[i],field_list[i],value_list[i])
    exit_excel();
