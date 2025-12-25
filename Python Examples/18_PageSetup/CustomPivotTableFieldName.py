from spire.xls.common import *
from spire.xls import *

inputFile = "Data/CustomPivotTableFieldName.xlsx"
outputFile = "CustomPivotTableFieldName-out.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the Excel document from disk
workbook.LoadFromFile(inputFile)

# Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]

# Access the first pivot table in the worksheet
pt = sheet.PivotTables.get_Item(0)

# Set a custom name for the column field
pt.ColumnFields[0].CustomName = "custom_colName"

# Set a custom name for the data field
pt.DataFields[0].CustomName = "custom_DataName"

# Calculate the pivot table data
pt.CalculateData()

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()