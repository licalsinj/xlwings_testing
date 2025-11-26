"""Trying to call a macro"""
import xlwings as xw
from pathlib import Path

p = Path(__file__).parent.parent
print(p)


wb = xw.Book(p._str + "/Workbooks/workbook_with_macro.xlsm")

# xlwings macro call: 
    # my_macro_name = wb.macro("Module.MyMacroName") # reference it like this
    # my_macro_name() # Call it like this. If you've got a param give it here
test_macro = wb.macro("Module1.Test_Module")
test_macro() 
