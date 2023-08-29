from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
inputimage = "./Demos/Data/Logo.png"
outputFile = "AddOvalShape.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add oval shape1
ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100)
ovalShape1.Line.Weight = 0
#Fill shape with solid color
ovalShape1.Fill.FillType = ShapeFillType.SolidColor
ovalShape1.Fill.ForeColor = Color.get_DarkCyan()
#Add oval shape2
ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100)
ovalShape2.Line.Weight = 1
#Fill shape with picture
ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid
ovalShape2.Fill.CustomPicture(inputimage)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
