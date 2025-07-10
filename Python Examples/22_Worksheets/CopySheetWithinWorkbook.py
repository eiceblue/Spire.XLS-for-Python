from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "CopySheetWithinWorkbook.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first and the second worksheets.
sheet = workbook.Worksheets[0]
sheet1 = workbook.Worksheets.Add("MySheet")
sourceRange = sheet.AllocatedRange
#Copy the first worksheet to the second one.
sheet.Copy(sourceRange, sheet1, sheet.FirstRow, sheet.FirstColumn, True)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
