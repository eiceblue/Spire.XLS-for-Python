from spire.xls.common import *
from spire.xls import *


outputFile = "CreateAnExcelWithOneSheet.xlsx"

workbook = Workbook()
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
for row in range(1, 100):
    for col in range(1, 31):
        sheet.Range[row,col].Text = str(row) + "," + str(col)

workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


