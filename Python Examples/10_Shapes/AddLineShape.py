from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "AddLineShape.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add shape line1
line1 = sheet.Lines.AddLine(10, 2, 200, 1, LineShapeType.Line)
#Set dash style type
line1.DashStyle = ShapeDashLineStyleType.Solid
#Set color
line1.Color = Color.get_CadetBlue()
#Set weight
line1.Weight = 2
#Set end arrow style type
line1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
#Add shape line2
line2 = sheet.Lines.AddLine(12, 2, 200, 1, LineShapeType.CurveLine)
line2.DashStyle = ShapeDashLineStyleType.Dotted
line2.Color = Color.get_OrangeRed()
line2.Weight = 2
#Add shape line3
line3 = sheet.Lines.AddLine(14, 2, 200, 1, LineShapeType.ElbowLine)
line3.DashStyle = ShapeDashLineStyleType.DashDotDot
line3.Color = Color.get_Purple()
line3.Weight = 2
#Add shape line4
line4 = sheet.Lines.AddLine(16, 2, 200, 1, LineShapeType.LineInv)
line4.DashStyle = ShapeDashLineStyleType.Dashed
line4.Color = Color.get_Green()
line4.Weight = 2
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

