from spire.xls.common import *
from spire.xls import *

inputFile = "Data/ExtractOleObjectName.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first sheet
sheet = workbook.Worksheets[0]

information = ""

for ole in sheet.OleObjects:
    # Obtain the original name of the document of the OLE object
    ole_name = ole.OriginName
    information += ole_name + "\r\n"

print(information)

# Dispose of the workbook object to release resources.
workbook.Dispose()
