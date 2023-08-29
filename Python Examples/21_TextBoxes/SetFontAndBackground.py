from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "SetFontAndBackground.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the textbox which will be edited.
shape = sheet.TextBoxes[0]
#Set the font and background color for the textbox.
#Set font.
font = workbook.CreateFont()
#font.IsStrikethrough = true
font.FontName = "Century Gothic"
font.Size = 10
font.IsBold = True
font.Color = Color.get_Blue()
rto = shape.RichText
rt = RichText(rto)
rt.SetFont(0, len(shape.Text) - 1, font)
#Set background color
shape.Fill.FillType = ShapeFillType.SolidColor
shape.Fill.ForeKnownColor = ExcelColors.BlueGray
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

