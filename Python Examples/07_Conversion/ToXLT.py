from spire.xls.common import *
from spire.xls import *

inputFile = "Data/ConversionTemplate.xlsx"
outputFile = "ToXLT.xlt"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Save the workbook to the specified output file with the specified file format (FileFormat::XLT).
workbook.SaveToFile(outputFile, FileFormat.XLT)

# Dispose of the workbook object to release resources.
workbook.Dispose()
