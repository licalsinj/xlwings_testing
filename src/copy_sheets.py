"""Wanted to test if I can set a worksheet to another and effectively copy the data over"""
import xlwings as xw

# create two workbooks
wb1 = xw.Book()
wb2 = xw.Book()

# add a sheet called "Main" to both
wb1.sheets.add("Main")
# wb2.sheets.add("Main")

# populate 1 with data
wb1.sheets["Main"].range('A2:E20').value = 100

# set another one's sheet equivalent to the others
wb1.sheets["Main"].api.Copy(Before=wb2.sheets[0].api)

# go check if they copied over 