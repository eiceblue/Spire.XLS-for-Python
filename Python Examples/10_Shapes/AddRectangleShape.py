from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "AddRectangleShape.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add rectangle shape 1------Rect
rect1 = sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect)
rect1.Line.Weight = 1
#Fill shape with solid color
rect1.Fill.FillType = ShapeFillType.SolidColor
rect1.Fill.ForeColor = Color.get_DarkGreen()
#Add rectangle shape 2------RoundRect
rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect)
rect2.Line.Weight = 1
rect2.Fill.FillType = ShapeFillType.SolidColor
rect2.Fill.ForeColor = Color.get_DarkCyan()
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
