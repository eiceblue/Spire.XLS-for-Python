from spire.xls.common import *
from spire.xls import *


for n in range(0, 50):
    workbook = Workbook()
    workbook.CreateEmptySheets(5)
    for i in range(0, 5):
        sheet = workbook.Worksheets[i]
        sheet.Name = "Sheet" + str(i)
        for row in range(1, 15):
            for col in range(1, 5):
                sheet.Range[row,col].Text = "row" + str(row) + " col" + str(col)

    workbook.SaveToFile("Workbook" + str(n) + ".xlsx", ExcelVersion.Version2010)
    workbook.Dispose()

