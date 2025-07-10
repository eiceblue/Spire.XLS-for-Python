from spire.xls import *
from spire.xls.common import *


outputFile = "LockSpecificCellInNewExcel.xlsx"

#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the rows in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock specific cell in the worksheet.
sheet.Range["A1"].Text = "Locked"
sheet.Range["A1"].Style.Locked = True
#Lock specific cell range in the worksheet.
sheet.Range["C1:E3"].Text = "Locked"
sheet.Range["C1:E3"].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

