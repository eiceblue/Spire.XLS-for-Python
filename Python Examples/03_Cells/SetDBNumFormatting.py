from spire.xls import *
from spire.common import *


outputFile = "SetDBNumFormatting.xlsx"

#Create a workbook
workbook = Workbook()
workbook.CreateEmptySheets(1)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set value for cells
sheet.Range["A1"].Value2 = Int32(123)
sheet.Range["A2"].Value2 = Int32(456)
sheet.Range["A3"].Value2 = Int32(789)
#Get the cell range
range = sheet.Range["A1:A3"]
#Set the DB num format
range.NumberFormat = "[DBNum2][$-804]General"
#Auto fit columns
range.AutoFitColumns()
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()



