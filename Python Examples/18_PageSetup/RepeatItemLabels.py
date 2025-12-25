from spire.xls.common import *
from spire.xls import *


inputFile = "Data/PivotTableExample.xlsx"
outputFile = "SetRepeatAllItemLabels-out.xlsx"


# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the sheet where the pivot table is located.
sheet = workbook.Worksheets["PivotTable"]

# Traverse the PivotTable and set options.
for pt in sheet.PivotTables:
    pt.Options.RepeatAllItemLabels(True)

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()
