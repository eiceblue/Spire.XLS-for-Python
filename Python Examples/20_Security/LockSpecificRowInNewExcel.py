from spire.xls import *
from spire.common import *


outputFile = "LockSpecificRowInNewExcel.xlsx"

#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the rows in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock the third row in the worksheet.
sheet.Rows[2].Text = "Locked"
sheet.Rows[2].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

