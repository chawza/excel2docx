from docx import Document

from typing import List

class Testcase:
    _id: str
    category: str
    description: str
    manual_steps: str
    expected_result: str
    actual_result: str
    datetime_test: str
    result: str

class Scenario:
    testcases: List[Testcase]

class Story:
    title :str
    _id: str
    description: str
    scenario: list[Scenario]


    def convert() -> Document:
        pass

