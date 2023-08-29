from spire.xls import *
from spire.common import *


outputFile = "SetCommentFillColor.xlsx"

#Create a workbook
workbook = Workbook()
#Get the default first worksheet
sheet = workbook.Worksheets[0]
#Create Excel font
font = workbook.CreateFont()
font.FontName = "Arial"
font.Size = 11
font.KnownColor = ExcelColors.Orange
#Add the comment
range = sheet.Range["A1"]
range.Comment.Text = "This is a comment"
range.Comment.RichText.SetFont(0, (len(range.Comment.Text) - 1), font)
#Set comment Color
range.Comment.Fill.FillType = ShapeFillType.SolidColor
range.Comment.Fill.ForeColor = Color.get_SkyBlue()
range.Comment.Visible = True
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
