from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "HideOrShowRowColumnHeaders.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Hide the headers of rows and columns
sheet.RowColumnHeadersVisible = False
#Show the headers of rows and columns
#sheet.RowColumnHeadersVisible = true
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

