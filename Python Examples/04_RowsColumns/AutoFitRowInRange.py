from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AutoFitSample.xlsx"
outputFile = "AutoFitRowInRange.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
# Autofit the second row of the worksheet
sheet.AutoFitRow(2, 1, 2, False)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

