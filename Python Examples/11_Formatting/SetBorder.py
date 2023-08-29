from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetBorder.xlsx"
outputFile = "SetBorder.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the cell range where you want to apply border style
cr = sheet.Range[sheet.FirstRow,sheet.FirstColumn,sheet.LastRow,sheet.LastColumn]
#Apply border style 
cr.Borders.LineStyle = LineStyleType.Double
cr.Borders[BordersLineType.DiagonalDown].LineStyle = LineStyleType.none
cr.Borders[BordersLineType.DiagonalUp].LineStyle = LineStyleType.none
cr.Borders.Color = Color.get_CadetBlue()
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

