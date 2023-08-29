from spire.xls import *
from spire.common import *


outputFile = "ApplyDataBarsToCellRange.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Insert data to cell range from A1 to C4.
sheet.Range["A1"].NumberValue = 582
sheet.Range["A2"].NumberValue = 234
sheet.Range["A3"].NumberValue = 314
sheet.Range["A4"].NumberValue = 50
sheet.Range["B1"].NumberValue = 150
sheet.Range["B2"].NumberValue = 894
sheet.Range["B3"].NumberValue = 560
sheet.Range["B4"].NumberValue = 900
sheet.Range["C1"].NumberValue = 134
sheet.Range["C2"].NumberValue = 700
sheet.Range["C3"].NumberValue = 920
sheet.Range["C4"].NumberValue = 450
sheet.AllocatedRange.RowHeight = 15
sheet.AllocatedRange.ColumnWidth = 17
#Add data bars.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(sheet.AllocatedRange)
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.DataBar
format.DataBar.BarColor = Color.get_CadetBlue()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

