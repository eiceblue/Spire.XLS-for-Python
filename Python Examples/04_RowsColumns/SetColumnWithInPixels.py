from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "SetColumnWithInPixels.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the width of the third column to 400 pixels
sheet.SetColumnWidthInPixels(3, 400)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


