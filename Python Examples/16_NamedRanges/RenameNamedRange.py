from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "RenameNamedRange.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Rename the named range
workbook.NameRanges[0].Name = "RenameRange"
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

