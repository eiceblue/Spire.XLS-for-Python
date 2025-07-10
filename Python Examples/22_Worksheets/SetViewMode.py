from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SetViewMode.xlsx"
outputFile = "SetViewMode.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Set the view mode 
workbook.Worksheets[0].ViewMode = ViewMode.Preview
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

