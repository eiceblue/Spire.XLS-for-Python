from spire.xls import *

inputFile = "./Demos/Data/MarkDownFile.md"
outputFile = "MarkDownFileout.xlsx"

# Create a new workbook
workbook = Workbook()

# Load the document
workbook.LoadFromMarkdown(inputFile)

# Save the Markdown as Excel
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()