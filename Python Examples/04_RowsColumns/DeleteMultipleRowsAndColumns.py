from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CommonTemplate1.xlsx"
outputFile = "DeleteMultipleRowsAndColumns.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Delete 4 rows from the fifth row
sheet.DeleteRow(5, 4)
#Delete 2 columns from the second column
sheet.DeleteColumn(2, 2)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()



