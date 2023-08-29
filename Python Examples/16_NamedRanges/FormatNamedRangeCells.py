from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/AllNamedRanges.xlsx"
outputFile = "FormatNamedRangeCells.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get specific named range by index
NamedRange = workbook.NameRanges[0]
#Get the cell range of the named range
range = NamedRange.RefersToRange
#Set color for the range
range.Style.Color = Color.get_Yellow()
#Set the font as bold
range.Style.Font.IsBold = True
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

