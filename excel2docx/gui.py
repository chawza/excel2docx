# author: nabeelkahlil403@gmail.com
# TODO: creates TK() class instance

import tkinter as tk
import os
from tkinter.filedialog import askopenfile, askdirectory
from tkinter.messagebox import showerror, showinfo, showwarning
from app import convert, rename_tc_filename
from exceptions import ReadWorksheetError
from openpyxl import load_workbook

DEFAULT_SOURCE_FILE_PATH = 'Source file path'
DEFAULT_FOLDER_PATH = os.path.join(os.getcwd(), 'results')

def askfor_source_file():
    file = askopenfile(
        filetypes = [('Excel files', '.xlsx .xls')],
        initialdir = os.getcwd()
    )
    if file is None:
        return

    global source_label_value
    source_label_value.config(text=os.path.abspath(file.name))

def askfor_target_directory():
    global target_label_value
    folder_location = askdirectory()

    if folder_location is None:
        return

    folder_location = os.path.abspath(folder_location)
    target_label_value.config(text=folder_location)

def create_result_folder_if_not_exist():
    package_dir = os.getcwd()
    file_list = os.listdir(package_dir)
    
    if 'results' not in file_list:
        os.makedirs(os.path.join(os.getcwd(), 'results'))

def process_file():
    global source_label_value
    global target_label_value

    source_file_path = source_label_value.cget('text')
    target_directory = target_label_value.cget('text')

    if source_file_path == DEFAULT_SOURCE_FILE_PATH:
        showwarning(title='Invalid source', message='you should specify your Excel location')
        return
    
    source_filename = os.path.basename(source_file_path) # convert path to excel filename
    source_filename = source_filename.split('.')[0]   # strip of excel extentions

    target_filename = source_filename + '.docx'
    target_filepath = None

    if 'TC' == target_filename[0:2]:
        target_filename = rename_tc_filename(target_filename)

    workbook = load_workbook(filename=source_file_path, read_only=True)
    try:
        doc = convert(workbook)
    except ReadWorksheetError as err:
        showerror(title="Worksheet Error", message=err.message)

    if target_directory == DEFAULT_FOLDER_PATH:
        create_result_folder_if_not_exist()
        target_filepath = os.path.join(DEFAULT_FOLDER_PATH, target_filename)
    else:
        target_filepath = os.path.join(target_directory, target_filename)

    doc.save(target_filepath)

    showinfo(title='saved docuemnt', message='document has been converted!')

app = tk.Tk()

frame = tk.Frame(master=app, width=400, height=400)
frame.grid()

title_label = tk.Label(master=frame, text="Excel2Docx", font=('', 30)).grid(columnspan=3, row=0)

source_label = tk.Label(master=frame, text='Source\t:')
source_label_value = tk.Label(master=frame, text=DEFAULT_SOURCE_FILE_PATH)
browse_source_button = tk.Button(master=frame, text='Source', command=askfor_source_file)
source_label.grid(column=0, row=1)
source_label_value.grid(column=1, row=1)
browse_source_button.grid(column=2, row=1)

target_label = tk.Label(master=frame, text='location\t:')
target_label_value = tk.Label(master=frame, text=DEFAULT_FOLDER_PATH)
browse_target_button = tk.Button(master=frame, text='Target', command=askfor_target_directory)
target_label.grid(column=0, row=2)
target_label_value.grid(column=1, row=2)
browse_target_button.grid(column=2, row=2)

process_button = tk.Button(master=frame, text="PROCESS", command=process_file)
process_button.grid(columnspan=3, row=3)

app.mainloop()