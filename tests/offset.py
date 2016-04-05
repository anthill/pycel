import os
import sys

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../src')
sys.path.insert(0, path)

# We will choose our wrapper with os compatibility
try:
    import win32com.client
    import pythoncom
    from pycel.excelwrapper import ExcelComWrapper as ExcelWrapperImpl
except:
    print "Can\'t import win32com -> switch from Com to Openpyxl wrapping implementation"
    from pycel.excelwrapper import ExcelOpxWrapper as ExcelWrapperImpl

# RUN AT THE ROOT LEVEL
excel = ExcelWrapperImpl(os.path.join(dir, "../example/example.xlsx"))
excel.connect()

def simple_offset():
    formula = excel.get_formula(3,4)
    parsed = excel.OffsetParser.parseOffsets(formula, excel.get_sheet().title)
    assert parsed[0] == '=Sheet1!B2'


simple_offset()
