from spire.xls.common import *
from spire.xls import *

inputFile = "Data/Shape.xlsx"
outputFile = "Shape.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first worksheet.
sheet = workbook.Worksheets[0]

# Get all shapes
shapelist = SaveShapeTypeOption()
shapelist.SaveAll = True

# Save shapes to images
images = sheet.SaveShapesToImage(shapelist)
for i in range(len(images)):
    images[i].Save("image_{0}.png".format(i))
    
workbook.Dispose()
