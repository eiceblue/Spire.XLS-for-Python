from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTable_1.xlsx"
outputFile = "CreateChartBasedOnPivotTable.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets[0]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
workbook.Worksheets[1].Charts.Add(ExcelChartType.BarClustered, pt)
#Save the Excel file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

