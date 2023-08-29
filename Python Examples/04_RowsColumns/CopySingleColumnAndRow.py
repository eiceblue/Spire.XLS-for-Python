from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "CopySingleColumnAndRow.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Specify a destination range to copy one column 
columnCells = sheet1.Range["G1:G19"]
#Copy the second column to destination range 
sheet1.Columns[1].Copy(columnCells)
#Specify a destination range to copy one row 
rowCells = sheet1.Range["A21:E21"]
#Copy the first row to destination range 
sheet1.Rows[0].Copy(rowCells)
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
