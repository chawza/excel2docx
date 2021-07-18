#By: Nabeel Kahlil Maulana

import sys
import os
from typing import List

from openpyxl import load_workbook
from docx import Document
from openpyxl.worksheet.worksheet import Worksheet

TABLE_STYLE_NAME = 'Light Grid Accent 6' # https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html
FIRST_ROW_TO_SCAN = 13

class TESTCASE_INDEX:
    ID = 0
    NAME = 1
    DESC = 3
    EXPECT_RESULT = 5
    ACTUAL_RESULT = 6

def read_testcases(ws: Worksheet) -> dict:
    testcases = []
    sliced_ws = ws[FIRST_ROW_TO_SCAN: ws.max_row-1]

    for index, row in enumerate(sliced_ws):
        if row[0].value is None:                # skip blank space between testcases.
            next_row = sliced_ws[index+1]   
            if next_row[0].value is None:       # last testcase followed with trail of blank spaces
                break
            continue

        testcases.append({
            'ID'                : row[TESTCASE_INDEX.ID].value,
            'name'              : row[TESTCASE_INDEX.NAME].value,
            'description'       : row[TESTCASE_INDEX.DESC].value,
            'expected_result'   : row[TESTCASE_INDEX.EXPECT_RESULT].value,
            'actual_result'     : row[TESTCASE_INDEX.ACTUAL_RESULT].value
        })

    return testcases

def get_story_data(ws: Worksheet) -> dict:
    story = {
        'title': '',
        'description': '',
        'story_id': ''
    }
    ws = ws[1:ws.max_row]
    for row in ws:
        story[row[0].value] = row[1].value
    return story

def get_uac(ws: Worksheet) -> list:
    uacs = []
    for row in ws:
        uac, no = row[0].value, int(row[1].value)
        if uac is not None and no >= 0:
            uacs.append({
                'uac': uac,
                'no': no
            })
    return uacs

def write_document(story: dict, testcases: list, uacs: List[dict]) -> Document:
    doc = Document()
    
    doc.add_heading(story['title'], 0)
    story_id_paragraph = doc.add_paragraph('')
    story_id_paragraph.add_run(story['story_id']).bold = True
    doc.add_paragraph(story['description'])

    for index, testcase in enumerate(testcases):
        if len(uacs) > 0 and uacs[0]['no'] >= index + 1:
            doc.add_heading(uacs[0]['uac'], 1)
            uacs.pop()
        
        doc.add_heading(f'Testcase{index+1}-{story["story_id"]}-{testcase["name"]}', 2)
        doc.add_paragraph(testcase['description'])
        doc.add_paragraph("[ADD CONTENT HERE]")

        table = doc.add_table(rows=2, cols=2)
        table.style = TABLE_STYLE_NAME
        table.cell(0,0).text = 'Expected Results'
        table.cell(0,1).text = 'Actual Results'
        table.cell(1,0).text = testcase['expected_result'] or ""
        table.cell(1,1).text = testcase['actual_result'] or ""
    
    return doc

def windows_to_unix_path(path: str):
    if '\\' in path:
        return path.replace('\\', '/')
    return path

def convert(file_location):
    workbook = load_workbook(filename=file_location, read_only=True)

    story = get_story_data(workbook['story'])
    testcases = read_testcases(workbook['Sprint 1'])
    uacs = get_uac(workbook['uac'])

    doc = write_document(story, testcases, uacs)
    return doc

if __name__ == '__main__':
    argv = sys.argv

    # find the excel
    if len(argv) == 1:
        raise Exception("require argument for file location")
    file_location = windows_to_unix_path(argv[1])
    file_name = file_location.split('/')[-1]    # get full file name

    # determine where to export
    # TODO: check wheter the destination is directory location or the absolue file path
    if len(argv) >= 3:
        result_location = windows_to_unix_path(argv[2])
        result_file_name = result_location.split('/')[-1]
        if len(result_file_name.split('.')) < 2: # check wheter target location have the correct postfix (.docx)
            result_location += '.docx'
    else:
        result_location = './results'+ '/' + file_name.split('.')[0] + '.docx'
    
    # process
    doc = convert(file_location)

    # compile and save
    try:
        doc.save(result_location)
    except FileNotFoundError as e:
        if e.errno == 2:
            os.makedirs('./results')
            doc.save(result_location) 
        else:
            raise e
