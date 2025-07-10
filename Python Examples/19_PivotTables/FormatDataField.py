from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/FormatDataField.xlsx"
outputFile = "FormatDataField.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
# Access the PivotTable
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
# Access the data field.
pivotDataField = pt.DataFields[0]
# Set data display format
pivotDataField.ShowDataAs = PivotFieldFormatType.PercentageOfColumn
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

