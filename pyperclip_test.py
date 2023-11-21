import pyperclip


vba_file_path=r'C:\Users\Tom\Downloads\VBA script.txt'
checking_path=r"C:\Users\Tom\Downloads\checking.txt"

with open(checking_path,'r') as file:
    checking_text=file.read()
#print(checking_text)
print(type(checking_text))
#print(vba_script)
s=pyperclip.paste()
print(type(s))
#print(s)
if s==checking_text:
    print("same")

#print(pyperclip.paste())