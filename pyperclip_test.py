import pyperclip


vba_file_path=r'C:\Users\Tom\Downloads\VBA script.txt'
with open(vba_file_path,'r') as file:
    vba_script=file.read()
print(vba_script)

pyperclip.copy(vba_script)
print(pyperclip.paste())