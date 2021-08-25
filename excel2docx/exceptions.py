class ReadWorksheetError(Exception):
    """Cannot find Testcase worksheet"""
    def __init__(self, worksheet_name = None):
        self.worksheet_name = worksheet_name
        tc_name = f'"{self.worksheet_name}"' if worksheet_name else ''
        self.message = f"Cannot find Testcase {tc_name} Worksheet"

    def __str__(self):
        return self.message