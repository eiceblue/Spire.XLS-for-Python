from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SpireXls.png"
outputFile = "InsertShapesToExcelSheet.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add a triangle shape.
triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle)
#Fill the triangle with solid color.
triangle.Fill.ForeColor = Color.get_Yellow()
triangle.Fill.FillType = ShapeFillType.SolidColor
#Add a heart shape.
heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart)
#Fill the heart with gradient color.
heart.Fill.ForeColor = Color.get_Red()
heart.Fill.FillType = ShapeFillType.Gradient
#Add an arrow shape with default color.
arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow)
#Add a cloud shape.
cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud)
#Fill the cloud with custom picture
cloud.Fill.CustomPicture(Stream(inputFile), "SpireXls.png")
cloud.Fill.FillType = ShapeFillType.Picture
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

