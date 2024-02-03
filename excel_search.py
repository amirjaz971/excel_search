import os
import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askstring
import subprocess as sp
def search_in_excel(directory, search_value):
    string=''
    excel_files = [file for file in os.listdir(directory) if file.endswith('.xlsx')]
    for excel_file in excel_files:
        file_path = os.path.join(directory, excel_file)
        excel_data = pd.read_excel(file_path, sheet_name=None)
        for sheet_name, sheet_data in excel_data.items():
            if sheet_data.map(lambda x: str(x).lower() if pd.notna(x) else x).eq(str(search_value).lower()).any().any():
                #row,col=divmod((sheet_data==search_value).to_numpy().nonzero()[0][0],len(sheet_data.columns))
                string+=f"Found value: '{search_value}' in sheet: '{sheet_name}' of file: '{excel_file}'"+"\n"                 
    return string                 
if __name__=='__main__':
    input_file=askopenfilename()
    input_value=askstring('Search','Enter the value: ').lower().rstrip()
    directory_path=os.path.dirname(input_file)                
    log=search_in_excel(directory_path, input_value)
    sp.Popen(['notepad', "search.txt"]) 
    file=open('search.txt','w',encoding="utf-16-le")    
    if log=='':
        file.write(f'{input_value} not found!')  
    else:      
        file.write(log)