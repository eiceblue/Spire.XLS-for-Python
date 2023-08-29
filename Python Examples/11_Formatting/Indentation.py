from spire.xls import *
from spire.common import *


outputFile = "Indentation.xlsx"

#Create a workbook
workbook = Workbook()
#Add a new worksheet to the Excel object
sheet = workbook.Worksheets[0]
#Access the "B5" cell from the worksheet
cell = sheet.Range["B5"]
#Add some value to the "B5" cell
cell.Text = "Hello Spire!"
#Set the indentation level of the text (inside the cell) to 2
cell.Style.IndentLevel = 2
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
