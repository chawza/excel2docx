#By: Nabeel Kahlil Maulana

import os

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from .story import TestFile, SCENARIO_MODE

def windows_to_unix_path(path: str):
    if '/' in path:
        return path
    return path.replace('\\', '/')

def convert(workbook: Workbook, scenario_read_mode):
    testfile = TestFile(workbook)
    testfile.read_testfile(scenario_read_mode=scenario_read_mode)
    return testfile.write_document()

def rename_tc_filename(filename: str):
    filename = list(filename)
    filename[0:2] = 'SS'
    return ''.join(filename)

def save_document(doc, result_location):
    if 'results' not in os.listdir():
        os.makedirs(os.path.join(os.getcwd(),'results'))
    doc.save(result_location)

def command_line_app(source_location, target_directory, scenario_read_mode=SCENARIO_MODE.UAC_SHEET):
    file_name = os.path.basename(source_location)
    target_filename = file_name.split('.')[0] + '.docx'

    if 'TC' == target_filename[0:2]:
        target_filename = rename_tc_filename(target_filename)
    
    workbook = load_workbook(filename=source_location, read_only=False)
    doc = convert(workbook, scenario_read_mode)

    target_filepath = os.path.join(target_directory, target_filename)
    save_document(doc, target_filepath)

    print(f'file saved in {target_filepath}')