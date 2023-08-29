from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CommonTemplate1.xlsx"
outputFile = "UnhideRowsAndColumns.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Unhide the row
sheet.ShowRow(15)
#Unhide th column
sheet.ShowColumn(4)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


