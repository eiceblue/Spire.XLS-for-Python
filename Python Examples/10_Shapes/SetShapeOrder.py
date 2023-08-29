from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetShapeOrder.xlsx"
outputFile = "SetShapeOrder.xlsx"

wb = Workbook()
#Load an excel file
wb.LoadFromFile(inputFile)
#Bring the picture forward one level
wb.Worksheets[0].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringForward)
#Bring the image in fron of all other objects
wb.Worksheets[1].Pictures[0].ChangeLayer(ShapeLayerChangeType.BringToFront)
#Send the shape back one level
shape = wb.Worksheets[2].PrstGeomShapes[1] if isinstance(wb.Worksheets[2].PrstGeomShapes[1], XlsShape) else None
shape.ChangeLayer(ShapeLayerChangeType.SendBackward)
#Send the shape behind all other objects
shape = wb.Worksheets[3].PrstGeomShapes[1] if isinstance(wb.Worksheets[3].PrstGeomShapes[1], XlsShape) else None
shape.ChangeLayer(ShapeLayerChangeType.SendToBack)
#Save to file.
wb.SaveToFile(outputFile, ExcelVersion.Version2010)
wb.Dispose()

