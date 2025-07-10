from spire.xls import *
from spire.xls.common import *


outputFile = "WriteRichText.xlsx"

workbook = Workbook()
sheet = workbook.Worksheets[0]

fontBold = workbook.CreateFont()
fontBold.IsBold = True

fontUnderline = workbook.CreateFont()
fontUnderline.Underline = FontUnderlineType.Single

fontItalic = workbook.CreateFont()
fontItalic.IsItalic = True

fontColor = workbook.CreateFont()
fontColor.KnownColor = ExcelColors.Green

richText = sheet.Range["B11"].RichText
richText.Text = "Bold and underlined and italic and colored text."
richText.SetFont(0, 3, fontBold)
richText.SetFont(9, 18, fontUnderline)
richText.SetFont(24, 29, fontItalic)
richText.SetFont(35, 41, fontColor)

workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()



