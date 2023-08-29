from spire.xls import *
from spire.common import *


outputFile = "SetConditionalFormatFormula.xlsx"

#Create a workbook
workbook = Workbook()
#Get the default first  worksheet
sheet = workbook.Worksheets[0]
#Add ConditionalFormat
xcfs = sheet.ConditionalFormats.Add()
#Define the range
xcfs.AddRange(sheet.Range["B5"])
#Add condition
format = xcfs.AddCondition()
format.FormatType = ConditionalFormatType.CellValue
#If greater than 1000
format.FirstFormula = "1000"
format.Operator = ComparisonOperatorType.Greater
format.BackColor = Color.get_Orange()
sheet.Range["B1"].NumberValue = 40
sheet.Range["B2"].NumberValue = 500
sheet.Range["B3"].NumberValue = 300
sheet.Range["B4"].NumberValue = 400
#Set a SUM formula for B5
sheet.Range["B5"].Formula = "=SUM(B1:B4)"
#Add text
sheet.Range["C5"].Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background."
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

