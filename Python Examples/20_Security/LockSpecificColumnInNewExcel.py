from spire.xls import *
from spire.xls.common import *


outputFile = "LockSpecificColumnInNewExcel.xlsx"

#Create a workbook.
workbook = Workbook()
#Create an empty worksheet.
workbook.CreateEmptySheet()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Loop through all the columns in the worksheet and unlock them.
for i in range(0,255):
    sheet.Rows[i].Style.Locked = False
#Lock the fourth column in the worksheet.
sheet.Columns[3].Text = "Locked"
sheet.Columns[3].Style.Locked = True
#Set the password.
sheet.Protect("123", SheetProtectionType.All)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
