from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/NamedRanges.xlsx"
outputFile = "NamedRanges.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Creating a named range
NamedRange = workbook.NameRanges.Add("NewNamedRange")
#Setting the range of the named range
NamedRange.RefersToRange = sheet.Range["A8:E12"]
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

