from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "ActivateWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the second worksheet from the workbook
sheet = workbook.Worksheets[1]
#Activate the sheet
sheet.Activate()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

