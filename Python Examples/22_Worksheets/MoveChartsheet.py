from spire.xls.common import *
from spire.xls import *

inputFile = "Data/MoveChartsheet.xlsx"
outputFile = "GoalSeek.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the Excel document from disk
workbook.LoadFromFile(inputFile)

# Get the first worksheet.
sheet = workbook.Worksheets[0]

# Move chart worksheets
workbook.Chartsheets[0].MoveSheet(2)
workbook.Chartsheets[0].MoveChartsheet(0)

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()