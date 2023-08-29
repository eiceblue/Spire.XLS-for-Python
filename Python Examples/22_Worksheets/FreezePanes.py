from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/FreezePanes.xlsx"
outputFile = "FreezePanes.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Freeze Top Row
sheet.FreezePanes(2, 1)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


