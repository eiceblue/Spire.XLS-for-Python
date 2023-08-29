from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "ScopedNamedRange.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add range name
namedRange = sheet.Names.Add("Range1")
#Define the range
namedRange.RefersToRange = sheet.Range["A1:D19"]
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

