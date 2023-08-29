from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "MergeNamedRangeCells.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the range of the named range
range = NamedRange.RefersToRange
#Merge cells
range.Merge()
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

