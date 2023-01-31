import csv
import os
import tkinter.filedialog
import xlwings as xw


ask_path = tkinter.filedialog.askdirectory(title="Select folder")
# Ask directory

os.chdir(ask_path)
# Active folder # Kích hoạt folder hiện hành

active_folder_name = os.getcwd()
# Get name of active folder

list_of_file = os.listdir()
# Get list of all file in active folder

new_file_excel = xw.Book()
# Open new Excel file

active_sheet = new_file_excel.sheets.active
# Get active sheet of Excel file
n = 2
def search_data_in_csv(filecsv_need_check):
    term = "Model"
    # Define the term which need to be search
    reader = csv.reader(open(filecsv_need_check,'r'))
    list_check = list(reader)
    for i in range(len(list_check)):
        if term in str(list_check[i]):
        # Check row include term
            active_sheet.range("B" + str(n)).value = list_check[i+1]
            active_sheet.range("A" + str(n)).value = filecsv_need_check
            print(list_check[i+1])
            break
for name_file in list_of_file:
    path = os.path.join(active_folder_name, name_file)
    # Define the path
    search_data_in_csv(path)
    n = n + 1