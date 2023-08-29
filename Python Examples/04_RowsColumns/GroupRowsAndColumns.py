from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/GroupRowsAndColumns.xls"
outputFile = "GroupRowsAndColumns.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Grouping rows
sheet.GroupByRows(1, 5, False)
#Grouping columns
sheet.GroupByColumns(1, 3, False)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

