from spire.xls.common import *
from spire.xls import *


inputFile = "./Demos/Data/CreateTable.xlsx"
outputFile = "CreateTable.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]

# Add a new List Object to the worksheet
sheet.ListObjects.Create("table", sheet.Range[1,1,19,5])
# Add Default Style to the table
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleLight9
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

