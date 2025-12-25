from spire.xls.common import *
from spire.xls import *

inputFile = "Data/Logo.png"
outputFile = "ImageHeaderFooterOnFirstPage.xlsx"

# Create a workbook.
workbook = Workbook()

# Get the first worksheet.
sheet = workbook.Worksheets[0]

# Set value for the range
cell1 = sheet.Range["A1"]
cell1.Text="Hello World"
cell2 = sheet.Range["F30"]
cell2.Text="Hello World"
cell3 = sheet.Range["G150"]
cell3.Text="Hello World"

# Set the value to show the headers/footers for first page are different from the other pages.
sheet.PageSetup.DifferentFirst = 1

imageStream = Stream(inputFile)
sheet.PageSetup.SetFirstLeftHeaderImage(imageStream)
sheet.PageSetup.SetFirstLeftFooterImage(imageStream)
sheet.PageSetup.LeftHeader = "Demo of Spire.XLS"
sheet.PageSetup.LeftFooter = "Footer by Spire.XLS"

sheet.ViewMode = ViewMode.Layout

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
