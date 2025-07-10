from spire.xls import *
from spire.xls.common import *


outputFile = "SetShadowStyleForShape.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add an ellipse shape.
ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse)
#Set the shadow style for the ellipse.
ellipse.Shadow.Angle = 90
ellipse.Shadow.Distance = 10
ellipse.Shadow.Size = 150
ellipse.Shadow.Color = Color.get_Gray()
ellipse.Shadow.Blur = 30
ellipse.Shadow.Transparency = 1
ellipse.Shadow.HasCustomStyle = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

