from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "ApplyMultipleFontsInSingleCell.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Create a font object in workbook, setting the font color, size and type.
font1 = workbook.CreateFont()
font1.KnownColor = ExcelColors.LightBlue
font1.IsBold = True
font1.Size = 10
#Create another font object specifying its properties.
font2 = workbook.CreateFont()
font2.KnownColor = ExcelColors.Red
font2.IsBold = True
font2.IsItalic = True
font2.FontName = "Times New Roman"
font2.Size = 11
#Write a RichText string to the cell 'A1', and set the font for it.
richText = sheet.Range["H5"].RichText
richText.Text = "This document was created with Spire.XLS for python."
richText.SetFont(0, 29, font1)
richText.SetFont(31, 48, font2)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

