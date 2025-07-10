from spire.xls import *
from spire.xls.common import *


outputFile = "FormatARow.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Create a new style
style = workbook.Styles.Add("newStyle")
#Set the vertical alignment of the text
style.VerticalAlignment = VerticalAlignType.Center
#Set the horizontal alignment of the text
style.HorizontalAlignment = HorizontalAlignType.Center
#Set the font color of the text
style.Font.Color = Color.get_Blue()
#Shrink the text to fit in the cell
style.ShrinkToFit = True
#Set the bottom border color of the cell to OrangeRed
style.Borders[BordersLineType.EdgeBottom].Color = Color.get_OrangeRed()
#Set the bottom border type of the cell to Dotted
style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Dotted
#Apply the style to the second row
sheet.Rows[1].CellStyleName = style.Name
sheet.Rows[1].Text = "Test"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

