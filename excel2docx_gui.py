# author: nabeelkahlil403@gmail.com
# TODO: creates TK() class instance

import tkinter as tk
import os
from tkinter.filedialog import askopenfile, askdirectory
from tkinter.messagebox import showinfo, showwarning
from excel2docx import convert

DEFAULT_SOURCE_FILE_PATH = 'Source file path'
DEFAULT_FOLDER_PATH = os.path.join(os.path.abspath(os.path.dirname(os.sys.argv[0])), 'results')

def askfor_source_file():
    file = askopenfile(
        filetypes=[('Excel files', '.xlsx .xls')],
        initialdir='./'
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
    package_dir = os.path.abspath(os.path.dirname(__file__))
    file_list = os.listdir(package_dir)
    
    if 'results' not in file_list:
        os.makedirs('./results')

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

    if 'TC' == target_filename.split('_')[0]:
        target_filename = list(target_filename)
        target_filename[0] = 'S'
        target_filename[1] = 'S'
        target_filename = ''.join(target_filename)

    target_filepath = None

    doc = convert(source_file_path)

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
browse_source_button = tk.Button(master=frame, text='browse source', command=askfor_source_file)
source_label.grid(column=0, row=1)
source_label_value.grid(column=1, row=1)
browse_source_button.grid(column=2, row=1)

target_label = tk.Label(master=frame, text='location\t:')
target_label_value = tk.Label(master=frame, text=DEFAULT_FOLDER_PATH)
browse_target_button = tk.Button(master=frame, text='browse target', command=askfor_target_directory)
target_label.grid(column=0, row=2)
target_label_value.grid(column=1, row=2)
browse_target_button.grid(column=2, row=2)

process_button = tk.Button(master=frame, text="PROCESS", command=process_file)
process_button.grid(columnspan=3, row=3)

app.mainloop()