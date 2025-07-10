from spire.xls.common import *
from spire.xls import *


outputFile = "CreateAnExcelWithFiveSheet.xlsx"

workbook = Workbook()
workbook.CreateEmptySheets(5)
for i in range(0, 5):
    sheet = workbook.Worksheets[i]
    sheet.Name = "Sheet" + str(i)
    for row in range(1, 151):
        for col in range(1, 51):
            sheet.Range[row,col].Text = "row" + str(row) + " col" + str(col)

workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
