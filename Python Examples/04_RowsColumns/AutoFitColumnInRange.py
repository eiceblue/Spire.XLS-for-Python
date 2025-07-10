from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/AutoFitSample.xlsx"
outputFile = "AutoFitColumnInRange.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Autofit the Column of the worksheet
sheet.AutoFitColumn(2, 2, 5)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

