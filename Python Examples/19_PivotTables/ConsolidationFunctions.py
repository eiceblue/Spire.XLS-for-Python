from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "ConsolidationFunctions.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#Apply Average consolidation function to first data field
pt.DataFields[0].Subtotal = SubtotalTypes.Average
#Apply Max consolidation function to second data field
pt.DataFields[1].Subtotal = SubtotalTypes.Max
pt.CalculateData()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


