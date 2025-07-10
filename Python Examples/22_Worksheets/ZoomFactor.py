from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ZoomFactor.xlsx"
outputFile = "ZoomFactor.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the zoom factor of the sheet to 85
sheet.Zoom = 85
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

