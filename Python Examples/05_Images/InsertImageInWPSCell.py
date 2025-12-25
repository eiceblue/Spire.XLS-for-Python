from spire.xls.common import *
from spire.xls import *

inputFile = "Data/Logo.png"
outputFile = "InsertImageInWPSCell.xlsx"

# Create a workbook.
workbook = Workbook()

# Get the first worksheet
sheet = workbook.Worksheets[0]

image = Stream(inputFile)

# Add an image in a cell range
sheet.Range["D1"].InsertOrUpdateCellImage(image,True)

# Save the workbook to the specified output file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

# Dispose of the workbook object to release resources.
workbook.Dispose()
