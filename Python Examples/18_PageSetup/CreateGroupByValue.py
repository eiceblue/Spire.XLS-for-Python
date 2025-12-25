from spire.xls.common import *
from spire.xls import *

inputFile = "Data/CreateGroupByValue.xlsx"
outputFile = "CreateGroupByValue-out.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first sheet
sheet = workbook.Worksheets[0]

# Access the first pivot table in the worksheet
pt = sheet.PivotTables.get_Item(0)

pivotField = pt.PivotFields["number"]
pivotField.CreateGroup(3000, 3800, 1)
pt.CalculateData()

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

# Dispose of the workbook object to release resources.
workbook.Dispose()
