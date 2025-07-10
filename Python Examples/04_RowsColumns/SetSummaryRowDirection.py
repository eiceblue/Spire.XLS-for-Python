from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "SetSummaryRowDirection.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Group rows
sheet.GroupByRows(1, 4, True)
#Set summary rows above details
sheet.PageSetup.IsSummaryRowBelow = False
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


