from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ShapeToImage.xlsx"
outputFile = "ShapeToImage.png"

wb = Workbook()
#Load an excel file
wb.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = wb.Worksheets[0]
#Get the first shape from the first worksheet
shape = sheet1.PrstGeomShapes[0] if isinstance(sheet1.PrstGeomShapes[0], XlsShape) else None
#Save the shape to a image
img = shape.SaveToImage()
img.Save(outputFile)
wb.Dispose()

