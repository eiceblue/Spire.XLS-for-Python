from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample3.xlsx"
outputFile = "HideOrShowWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Hide the sheet named "Sheet1"
workbook.Worksheets["Sheet1"].Visibility = WorksheetVisibility.Hidden
#Show the second sheet
workbook.Worksheets[1].Visibility = WorksheetVisibility.Visible
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
