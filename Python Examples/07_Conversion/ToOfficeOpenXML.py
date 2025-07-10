from spire.xls import *
from spire.xls.common import *


outputFile = "ToOfficeOpenXML.xml"

workbook = Workbook()
sheet = workbook.Worksheets[0]
sheet.Range["A1"].Text = "Hello World"
sheet.Range["B1"].Style.KnownColor = ExcelColors.Gray25Percent
sheet.Range["C1"].Style.KnownColor = ExcelColors.Gold
workbook.SaveAsXml(outputFile)
workbook.Dispose()



