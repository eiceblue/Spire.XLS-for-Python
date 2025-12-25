from spire.xls.common import *
from spire.xls import *

inputFile = "Data/CreateTable.xlsx"
outputFile = "ReplaceFont_out.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first worksheet.
sheet = workbook.Worksheets[0]

newStyle = workbook.Styles.Add("newStyle")
newStyle.Font.FontName = "Arial Black"
newStyle.Font.Size = 14

cellRange = sheet.Range["D9"]
oldStyle = cellRange.Style
sheet.ReplaceAll("North America", oldStyle, "America", newStyle)

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
