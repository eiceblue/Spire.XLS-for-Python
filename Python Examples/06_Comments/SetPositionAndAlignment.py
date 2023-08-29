from spire.xls import *
from spire.common import *


outputFile = "SetPositionAndAlignment.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Set two font styles which will be used in comments
font1 = workbook.CreateFont()
font1.FontName = "Calibri"
font1.Color = Color.get_Firebrick()
font1.IsBold = True
font1.Size = 12
font2 = workbook.CreateFont()
font2.FontName = "Calibri"
font2.Color = Color.get_Blue()
font2.Size = 12
font2.IsBold = True
#Add comment 1 and set its size, text, position and alignment
sheet.Range["G5"].Text = "Spire.XLS"
Comment1 = sheet.Range["G5"].Comment
Comment1.IsVisible = True
Comment1.Height = 150
Comment1.Width = 300
Comment1.RichText.Text = "Spire.XLS for .Net:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. "
Comment1.RichText.SetFont(0, 19, font1)
Comment1.TextRotation = TextRotationType.LeftToRight
#Set the position of Comment
Comment1.Top = 20
Comment1.Left = 40
#Set the alignment of text in Comment
Comment1.VAlignment = CommentVAlignType.Center
Comment1.HAlignment = CommentHAlignType.Justified
#Add comment2 and set its size, text, position and alignment for comparison
sheet.Range["D14"].Text = "E-iceblue"
Comment2 = sheet.Range["D14"].Comment
Comment2.IsVisible = True
Comment2.Height = 150
Comment2.Width = 300
Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents."
Comment2.TextRotation = TextRotationType.LeftToRight
Comment2.RichText.SetFont(0, 16, font2)
#Set the position of Comment
Comment2.Top = 170
Comment2.Left = 450
#Set the alignment of text in Comment
Comment2.VAlignment = CommentVAlignType.Top
Comment2.HAlignment = CommentHAlignType.Justified
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

