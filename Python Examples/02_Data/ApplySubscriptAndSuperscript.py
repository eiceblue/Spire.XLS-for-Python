from spire.xls.common import *
from spire.xls import *


outputFile = "ApplySubscriptAndSuperscript.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]
sheet.Range["B2"].Text = "This is an example of Subscript:"
sheet.Range["D2"].Text = "This is an example of Superscript:"

#Set the rtf value of "B3" to "R100-0.06".
range = sheet.Range["B3"]
range.RichText.Text = "R100-0.06"

#Create a font. Set the IsSubscript property of the font to "true".
font = workbook.CreateFont()
font.IsSubscript = True
font.Color = Color.get_Green()

#Set font for specified range of the text in "B3".
range.RichText.SetFont(4, 8, font)

#Set the rtf value of "D3" to "a2 + b2 = c2".
range = sheet.Range["D3"]
range.RichText.Text = "a2 + b2 = c2"

#Create a font. Set the IsSuperscript property of the font to "true".
font = workbook.CreateFont()
font.IsSuperscript = True

#Set font for specified range of the text in "D3".
range.RichText.SetFont(1, 1, font)
range.RichText.SetFont(6, 6, font)
range.RichText.SetFont(11, 11, font)

sheet.AllocatedRange.AutoFitColumns()
 
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()