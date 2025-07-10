from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/RemoveWorksheet.xlsx"
outputFile = "RemoveWorksheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
# Remove a worksheet by sheet index
workbook.Worksheets.RemoveAt(1)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

