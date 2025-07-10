from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "FormatCellsWithStyle.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Create a style
style = workbook.Styles.Add("newStyle")
#Set the shading color
style.Color = Color.get_DarkGray()
#Set the font color
style.Font.Color = Color.get_White()
#Set font name
style.Font.FontName = "Times New Roman"
#Set font size
style.Font.Size = 12
#Set bold for the font
style.Font.IsBold = True
#Set text rotation
style.Rotation = 45
#Set alignment
style.HorizontalAlignment = HorizontalAlignType.Center
style.VerticalAlignment = VerticalAlignType.Center
#Set the style for the specific range
workbook.Worksheets[0].Range["A1:J1"].CellStyleName = style.Name
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
