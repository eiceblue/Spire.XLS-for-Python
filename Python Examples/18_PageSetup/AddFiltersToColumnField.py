from spire.xls.common import *
from spire.xls import *

inputFile = "Data/PivotTableExample2.xlsx"
outputFile = "ColumnFieldFilter-out.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the sheet with pivot table
sheet = workbook.Worksheets["PivotTable"]

# Get the first pivot table
pt = sheet.PivotTables[0]

# Get the first column field
pt.ColumnFields.get_Item(0).AddLabelFilter(PivotLabelFilterType.Equal, String("Brasilia"), None)
pt.CalculateData()

# Save the workbook to the specified output file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)

# Dispose of the workbook object to release resources.
workbook.Dispose()
