from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/InsertRowsAndColumns.xls"
outputFile = "InsertRowsAndColumns.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
#Inserting a row into the worksheet 
worksheet.InsertRow(2)
#Inserting a column into the worksheet 
worksheet.InsertColumn(2)
#Inserting multiple rows into the worksheet
worksheet.InsertRow(5, 2)
#Inserting multiple columns into the worksheet
worksheet.InsertColumn(5, 2)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

