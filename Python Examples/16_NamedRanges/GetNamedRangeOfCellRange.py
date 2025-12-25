from spire.xls.common import *
from spire.xls import *

inputFile = "Data/AllNamedRanges.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first sheet
sheet = workbook.Worksheets[0]

# Get some cellRanges
cellRange = sheet.Range["A2:D2"]

# Get the named range object
result = cellRange.GetNamedRange()

# print the name of the named range object
print(result.Name)

# Dispose of the workbook object to release resources.
workbook.Dispose()
