from spire.xls.common import *
from spire.xls import *


outputFile = "AddScrollBarControl.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
#Set a value for range B10
sheet.Range["B10"].NumberValue = 1
sheet.Range["B10"].Style.Font.IsBold = True
#Add scroll bar control
scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20)
scrollBar.LinkedCell = sheet.Range["B10"]
scrollBar.Min = 1
scrollBar.Max = 150
scrollBar.IncrementalChange = 1
scrollBar.Display3DShading = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
