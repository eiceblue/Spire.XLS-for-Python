from spire.xls.common import *
from spire.xls import *

outputFile = "GroupShapes.xlsx"

# Create a workbook
workbook = Workbook()

# Get the first worksheet
sheet = workbook.Worksheets[0]

# Add shapes to the worksheet
shape1 = sheet.PrstGeomShapes.AddPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect)
shape2 = sheet.PrstGeomShapes.AddPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle)
groupShapeCollection = sheet.GetGroupShapeCollection() 
groupShapeCollection.Group([shape1, shape2])

# Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
