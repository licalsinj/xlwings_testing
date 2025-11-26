import xlwings as xw
import datetime as dt

# wb = xw.Book()
# wb.save('text2.xlsx') # saves the excel sheet at the top of this project's directory
# wb.close

wbLoad = xw.Book('test2.xlsx') # loads an excel sheet

worksheet1 = wbLoad.sheets[0] # can index the sheet manually or by name 'Sheet1'

print(worksheet1.name) 

worksheet1.range('A1').value = 'Hello, World'
worksheet1.range('A2:E20').value = 100

worksheet1.clear_contents() # clears all the contents of the sheet 

worksheet1.cells(1,1).value = 100 # access the cell by the row and column location
worksheet1.cells(1, "B").value = 200 # can also address columns by letter name 

# Tables:

worksheet1.range('A2').value = [['Col A', 'Col B'], [10, 20], [30, 40]] # this will create a table (not an excel table) with each row being  part of the 

worksheet1.range('A5').options(transpose=True).value = ["apple", "carrot", "eggs", "milk", "pizza!"] # will transpose a horizontal list to be vertical

print(worksheet1.range('A5').value) # reads the value of a single cell

print(worksheet1.range('A1').expand().value) # expand() will automatically expand the cells range to whatever's available
print(worksheet1.range('B1').options(expand='table').value) # expand table will select everything as a table

# convert dates

worksheet1.range('E3').value = '5/5/2025'
print(worksheet1.range('E3').value) # returns with midnight as time
print(worksheet1.range('E3').options(dates=dt.date).value) # returns without the time


worksheet1.range('H4:J4').value = ['Bob', None,'$200']
print(worksheet1.range('H4:J4').value) 
print(worksheet1.range('H4:J4').options(empty='NA').value) # replaces none value with NA

# to call a macro do this: 
    # xlwings macro call: 
    # my_macro_name = wb.macro("Module.MyMacroName") # reference it like this
    # my_macro_name() # Call it like this. If you've got a param give it here 
# if it's a macro work book do this: 
    # wb = xw.Book("MyBook.xls")
    # app = wb.app
    # macro_vba = app.macro("'PERSONAL.XLSB'!my_macro_name")
# I haven't tested either of the above but it's what I found so far