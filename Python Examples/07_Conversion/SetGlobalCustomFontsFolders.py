from spire.xls.common import *
from spire.xls import *

inputFontPath="\\font\\"
inputFile = "Data/ToPDFSample.xlsx"
outputFile = "SetGlobalCustomFontsFolders.pdf"

# Set global custom font folders
Workbook.SetGlobalCustomFontsFolders(inputFontPath)

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Save the workbook to the specified output file.
workbook.SaveToFile(outputFile, FileFormat.PDF)

# Dispose of the workbook object to release resources.
workbook.Dispose()
