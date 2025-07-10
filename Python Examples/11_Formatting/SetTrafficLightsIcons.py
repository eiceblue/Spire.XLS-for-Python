from spire.xls import *
from spire.xls.common import *


outputFile = "SetTrafficLightsIcons.xlsx"

#Create a workbook.
workbook = Workbook()
#Add a worksheet.
sheet = workbook.Worksheets[0]
#Add some data to the Excel sheet cell range and set the format for them.
sheet.Range["A1"].Text = "Traffic Lights"
sheet.Range["A2"].NumberValue = 0.95
sheet.Range["A2"].NumberFormat = "0%"
sheet.Range["A3"].NumberValue = 0.5
sheet.Range["A3"].NumberFormat = "0%"
sheet.Range["A4"].NumberValue = 0.1
sheet.Range["A4"].NumberFormat = "0%"
sheet.Range["A5"].NumberValue = 0.9
sheet.Range["A5"].NumberFormat = "0%"
sheet.Range["A6"].NumberValue = 0.7
sheet.Range["A6"].NumberFormat = "0%"
sheet.Range["A7"].NumberValue = 0.6
sheet.Range["A7"].NumberFormat = "0%"
#Set the height of row and width of column for Excel cell range.
sheet.AllocatedRange.RowHeight = 20
sheet.AllocatedRange.ColumnWidth = 25
#Add a conditional formatting.
conditional = sheet.ConditionalFormats.Add()
conditional.AddRange(sheet.AllocatedRange)
format1 = conditional.AddCondition()
#Add a conditional formatting of cell range and set its type to CellValue.
format1.FormatType = ConditionalFormatType.CellValue
format1.FirstFormula = "300"
format1.Operator = ComparisonOperatorType.Less
format1.FontColor = Color.get_Black()
format1.BackColor = Color.get_LightSkyBlue()
#Add a conditional formatting of cell range and set its type to IconSet.
conditional.AddRange(sheet.AllocatedRange)
format = conditional.AddCondition()
format.FormatType = ConditionalFormatType.IconSet
format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

