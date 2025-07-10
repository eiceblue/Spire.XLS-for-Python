from spire.xls.common import *
from spire.xls import *


outputFile = "CloneExcelFontStyle.xlsx"

 #Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Add the text to the Excel sheet cell range A1.
sheet.Range["A1"].Text = "Text1"

#Set A1 cell range's CellStyle.
style = workbook.Styles.Add("style")
style.Font.FontName = "Calibri"
style.Font.Color = Color.get_Red()
style.Font.Size = 12
style.Font.IsBold = True
style.Font.IsItalic = True
sheet.Range["A1"].CellStyleName = style.Name

#Clone the same style for B2 cell range.
csOrieign = style.clone()
sheet.Range["B2"].Text = "Text2"
sheet.Range["B2"].CellStyleName = csOrieign.Name

#Clone the same style for C3 cell range and then reset the font color for the text.
csGreen = style.clone()
csGreen.Font.Color = Color.get_Green()
sheet.Range["C3"].Text = "Text3"
sheet.Range["C3"].CellStyleName = csGreen.Name
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()