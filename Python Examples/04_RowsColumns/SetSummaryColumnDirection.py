from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "SetSummaryColumnDirection.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Group Columns
sheet.GroupByColumns(1, 4, True)
#Set summary columns to right of details
sheet.PageSetup.IsSummaryRowBelow = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


