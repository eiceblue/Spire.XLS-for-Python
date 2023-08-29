from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/TextAlign.xlsx"
outputFile = "TextAlign.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the vertical alignment to Top
sheet.Range["B1:C1"].Style.VerticalAlignment = VerticalAlignType.Top
#Set the vertical alignment to Center
sheet.Range["B2:C2"].Style.VerticalAlignment = VerticalAlignType.Center
#Set the vertical alignment of to Bottom
sheet.Range["B3:C3"].Style.VerticalAlignment = VerticalAlignType.Bottom
#Set the horizontal alignment to General
sheet.Range["B4:C4"].Style.HorizontalAlignment = HorizontalAlignType.General
#Set the horizontal alignment of to Left
sheet.Range["B5:C5"].Style.HorizontalAlignment = HorizontalAlignType.Left
#Set the horizontal alignment of to Center
sheet.Range["B6:C6"].Style.HorizontalAlignment = HorizontalAlignType.Center
#Set the horizontal alignment of to Right
sheet.Range["B7:C7"].Style.HorizontalAlignment = HorizontalAlignType.Right
#Set the rotation degree
sheet.Range["B8:C8"].Style.Rotation = 45
sheet.Range["B9:C9"].Style.Rotation = 90
#Set the row height of cell
sheet.Range["B8:C9"].RowHeight = 60
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

