from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WriteComment.xlsx"
outputFile = "WriteComment.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Creates font
font = workbook.CreateFont()
font.FontName = "Arial"
font.Size = 11
font.KnownColor = ExcelColors.Orange
fontBlue = workbook.CreateFont()
fontBlue.KnownColor = ExcelColors.LightBlue
fontGreen = workbook.CreateFont()
fontGreen.KnownColor = ExcelColors.LightGreen
range = sheet.Range["B11"]
range.Text = "Regular comment"
range.Comment.Text = "Regular comment"
range.AutoFitColumns()
#Regular comment
range = sheet.Range["B12"]
range.Text = "Rich text comment"
range.RichText.SetFont(0, 16, font)
range.AutoFitColumns()
#Rich text comment
range.Comment.RichText.Text = "Rich text comment"
range.Comment.RichText.SetFont(0, 4, fontGreen)
range.Comment.RichText.SetFont(5, 9, fontBlue)
workbook.SaveToFile(outputFile, ExcelVersion.Version2007)
workbook.Dispose()
