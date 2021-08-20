#By: Nabeel Kahlil Maulana

import sys
import os
from typing import List

from docx import Document
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

from Story import TestFile

TC_SHEET_NAME = 'Sprint 1'
STORY_SHEET_NAME = 'story'
UAC_SHEET_NAME = 'uac'
TABLE_STYLE_NAME = 'Light Grid Accent 6' # https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html
FIRST_ROW_TO_SCAN = 13

def windows_to_unix_path(path: str):
    if '/' in path:
        return path
    return path.replace('\\', '/')

def convert(workbook: Workbook, uac_sheet = False):
    tc_file = TestFile(workbook, uac_sheet=uac_sheet)
    return tc_file.get_docx()

def rename_tc_filename(filename: str):
    filename = list(filename)
    filename[0:2] = 'SS'
    return ''.join(filename)

def save_document(doc, result_location):
    if 'results' not in os.listdir():
        os.makedirs(os.join(os.getcwd(),'results'))
    doc.save(result_location)

def command_line_app(source_location, target_directory, uac_sheet: bool = False):
    file_name = os.path.basename(source_location)
    target_filename = file_name.split('.')[0] + '.docx'

    if 'TC' == target_filename[0:2]:
        target_filename = rename_tc_filename(target_filename)
    
    workbook = load_workbook(filename=source_location, read_only=True)
    doc = convert(workbook, uac_sheet=uac_sheet)

    target_filepath = os.path.join(target_directory, target_filename)
    save_document(doc, target_filepath)

    print(f'file saved in {target_filepath}')

if __name__ == '__main__':
    argv = sys.argv.copy()

    use_uac_sheet = False
    if '-uac' in argv:
        use_uac_sheet = True
        argv.remove('-uac')

    # TODO: recreate commandline engine
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

    command_line_app(source_location, target_directory, use_uac_sheet)
