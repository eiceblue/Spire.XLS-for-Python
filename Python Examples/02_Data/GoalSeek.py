from spire.xls.common import *
from spire.xls import *


outputFile = "GoalSeek.xlsx"

# Create a workbook.
workbook = Workbook()

# Get the first worksheet.
sheet = workbook.Worksheets[0]

# Set value for cell "A1"
sheet.Range["A1"].Value="100"

# Set formula for cell "A2"
target_cell = sheet.Range["A2"]
target_cell.Formula="=SUM(A1+B1)"

# Variable cell
guess_cell = sheet.Range["B1"]

goal_seek = GoalSeek()

result = goal_seek.TryCalculate(target_cell, 500.0, guess_cell)
result.Determine()

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
