from spire.xls.common import *
from spire.xls import *


outputFile = "HelloWorld.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
sheet = workbook.Worksheets.Add("MySheet")
sheet.Range["A1"].Text = "Hello World"
sheet.Range["A1"].AutoFitColumns()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

