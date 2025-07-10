from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SpireXls.png"
outputFile = "AlignPictureWithinCell.xlsx"
     
#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
sheet.Range["A1"].Text = "Align Picture Within A Cell:"
sheet.Range["A1"].Style.VerticalAlignment = VerticalAlignType.Top
picture = sheet.Pictures.Add(1, 1, inputFile)
#Adjust the column width and row height so that the cell can contain the picture.
sheet.Columns[0].ColumnWidth = 40
sheet.Rows[0].RowHeight = 200
#Vertically and horizontally align the image.
picture.LeftColumnOffset = 100
picture.TopRowOffset = 25
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
