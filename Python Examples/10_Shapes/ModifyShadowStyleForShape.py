from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_5.xlsx"
outputFile = "ModifyShadowStyleForShape.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the third shape from the worksheet.
shape = sheet.PrstGeomShapes[2]
#Set the shadow style for the shape.
shape.Shadow.Angle = 90
shape.Shadow.Transparency = 30
shape.Shadow.Distance = 10
shape.Shadow.Size = 130
shape.Shadow.Color = Color.get_Yellow()
shape.Shadow.Blur = 30
shape.Shadow.HasCustomStyle = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


