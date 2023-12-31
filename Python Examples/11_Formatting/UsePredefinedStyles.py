﻿from spire.xls import *
from spire.common import *


outputFile = "UsePredefinedStyles.xlsx"

workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a new style
style = workbook.Styles.Add("newStyle")
style.Font.FontName = "Calibri"
style.Font.IsBold = True
style.Font.Size = 15
style.Font.Color = Color.get_CornflowerBlue()
#Get "B5" cell
range = sheet.Range["B5"]
range.Text = "Welcome to use Spire.XLS"
range.CellStyleName = style.Name
range.AutoFitColumns()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
