from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "MoveWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Move worksheet
sheet.MoveWorksheet(2)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

