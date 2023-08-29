from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "RemoveNamedRange.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Remove the named range by index
workbook.NameRanges.RemoveAt(0)
#Remove the named range by name
workbook.NameRanges.Remove("NameRange2")
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

