from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_2.xlsx"
outputFile = "DeleteBlankRowsAndColumns.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Delete blank rows from the worksheet.
for i in range(sheet.Rows.Length - 1, -1, -1):
    if sheet.Rows[i].IsBlank:
        sheet.DeleteRow(i + 1)
#Delete blank columns from the worksheet.
for j in range(sheet.Columns.Length - 1, -1, -1):
    if sheet.Columns[j].IsBlank:
        sheet.DeleteColumn(j + 1)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

