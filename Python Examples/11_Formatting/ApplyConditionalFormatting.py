from spire.xls import *
from spire.common import *


outputFile = "ApplyConditionalFormatting.xlsx"

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
#Create conditional formatting rule.
xcfs1 = sheet.ConditionalFormats.Add()
xcfs1.AddRange(sheet.AllocatedRange)
format1 = xcfs1.AddCondition()
format1.FormatType = ConditionalFormatType.CellValue
format1.FirstFormula = "800"
format1.Operator = ComparisonOperatorType.Greater
format1.FontColor = Color.get_Red()
format1.BackColor = Color.get_LightSalmon()
#Create conditional formatting rule.
xcfs2 = sheet.ConditionalFormats.Add()
xcfs2.AddRange(sheet.AllocatedRange)
format2 = xcfs1.AddCondition()
format2.FormatType = ConditionalFormatType.CellValue
format2.FirstFormula = "300"
format2.Operator = ComparisonOperatorType.Less
format2.FontColor = Color.get_Green()
format2.BackColor = Color.get_LightBlue()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
