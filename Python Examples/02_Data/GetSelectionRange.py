from spire.xls.common import *
from spire.xls import *


inputFile = "Data/GetSelectionRange.xlsx"

# Create a workbook.
workbook = Workbook()

# Load the workbook from the specified input file.
workbook.LoadFromFile(inputFile)

# Get the first worksheet.
sheet = workbook.Worksheets[0]

ranges = sheet.GetActiveSelectionRange()
information = ""
for range in ranges:
    information += "RangeAddressLocal:" + range.RangeAddressLocal + "\r\n"
    information += "ColumnCount:" + str(range.ColumnCount) + "\r\n"
    information += "ColumnWidth:" + str(range.ColumnWidth) + "\r\n"
    information += "Column:" + str(range.Column) + "\r\n"
    information += "RowCount:" + str(range.RowCount) + "\r\n"
    information += "RowHeight:" + str(range.RowHeight) + "\r\n"
    information += "Row:" + str(range.Row) + "\r\n"

print(information)

workbook.Dispose()
