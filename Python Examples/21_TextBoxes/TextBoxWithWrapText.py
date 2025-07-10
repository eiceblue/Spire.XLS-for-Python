from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/TextBoxSampleB.xlsx"
outputFile = "TextBoxWithWrapText.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk          
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the text box
shape = sheet.TextBoxes[0] if isinstance(sheet.TextBoxes[0], XlsTextBoxShape) else None
#Set wrap text
shape.IsWrapText = True
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

