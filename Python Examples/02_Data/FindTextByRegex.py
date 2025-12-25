from spire.xls.common import *
from spire.xls import *

inputFile = "Data/FindTextByRegex.xlsx"
outputFile = "FindTextByRegex_out.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first worksheet.
sheet = workbook.Worksheets[0]

ranges = sheet.FindAllString(".*North.", False, False, True)

for range in ranges:
    # Highlight the cell range
    range.Style.Color = Color.get_Yellow()

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
