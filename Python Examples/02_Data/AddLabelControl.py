from spire.common import *
from spire.xls import *


outputFile = "AddLabelControl.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
label = sheet.LabelShapes.AddLabel(10, 2, 30, 200)
label.Text = "This is a Label Control"
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()