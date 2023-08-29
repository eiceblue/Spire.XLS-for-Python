from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/templateAz.xlsx"
outputFile = "GetStyleSetStyle.xlsx"

#Create a workbook
workbook = Workbook()
#Load a excel document
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get "B4" cell
range = sheet.Range["B4"]
#Get the style of cell
style = range.Style
style.Font.FontName = "Calibri"
style.Font.IsBold = True
style.Font.Size = 15
style.Font.Color = Color.get_CornflowerBlue()
range.Style = style
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

