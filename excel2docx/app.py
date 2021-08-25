#By: Nabeel Kahlil Maulana

import sys
import os
from typing import List

from openpyxl import load_workbook
from docx import Document
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

from story import TestFile

def windows_to_unix_path(path: str):
    if '/' in path:
        return path
    return path.replace('\\', '/')

def convert(workbook: Workbook, uac_sheet = True):
    testfile = TestFile(workbook)
    testfile.read_testfile(uac_sheet=uac_sheet)
    return testfile.write_document()

def rename_tc_filename(filename: str):
    filename = list(filename)
    filename[0:2] = 'SS'
    return ''.join(filename)

def save_document(doc, result_location):
    if 'results' not in os.listdir():
        os.makedirs(os.join(os.getcwd(),'results'))
    doc.save(result_location)

def command_line_app(source_location, target_directory, uac_sheet=True):
    file_name = os.path.basename(source_location)
    target_filename = file_name.split('.')[0] + '.docx'

    if 'TC' == target_filename[0:2]:
        target_filename = rename_tc_filename(target_filename)
    
    workbook = load_workbook(filename=source_location, read_only=True)
    doc = convert(workbook, uac_sheet)

    target_filepath = os.path.join(target_directory, target_filename)
    save_document(doc, target_filepath)

    print(f'file saved in {target_filepath}')

if __name__ == '__main__':
    argv = sys.argv

    is_uac_sheet = True
    if '--uac-tc' in argv:
        is_uac_sheet = False
        argv.remove('--uac-tc')

    try:
        source_location = argv[1]
        target_directory = argv[2]
    except IndexError:
        if len(argv) < 2:
            raise Exception("argument required: target source is not defined")
        if len(argv) < 3:
            target_directory = os.path.join(os.getcwd(), 'results')
    
    source_location = windows_to_unix_path(source_location)
    target_directory = windows_to_unix_path(target_directory)

    command_line_app(source_location, target_directory, uac_sheet=is_uac_sheet)
