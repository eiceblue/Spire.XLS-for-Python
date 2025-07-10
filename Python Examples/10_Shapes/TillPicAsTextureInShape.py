from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/TillPicAsTextureInShape.xlsx"
inputImage = "./Demos/Data/Logo.png"
outputFile = "TillPicAsTextureInShape.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the first shape
shape = sheet.PrstGeomShapes[0]
#Fill shape with texture
shape.Fill.FillType = ShapeFillType.Texture
#Custom texture with picture
shape.Fill.CustomTexture(inputImage)
#Tile pciture as texture 
shape.Fill.Tile = True
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


