from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/HideRowsAndColumns.xls"
outputFile = "HideRowsAndColumns.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
worksheet = workbook.Worksheets[0]
# Hiding the column of the worksheet
worksheet.HideColumn(2)
#Hiding the row of the worksheet
worksheet.HideRow(4)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

