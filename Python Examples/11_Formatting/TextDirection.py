from spire.xls import *
from spire.common import *


outputFile = "TextDirection.xlsx"

#Create a workbook
workbook = Workbook()
#Add a new worksheet to the Excel object
sheet = workbook.Worksheets[0]
#Access the "B5" cell from the worksheet
cell = sheet.Range["B5"]
#Add some value to the "B5" cell
cell.Text = "Hello Spire!"
#Set the reading order from right to left of the text in the "B5" cell
cell.Style.ReadingOrder = ReadingOrderType.RightToLeft
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

