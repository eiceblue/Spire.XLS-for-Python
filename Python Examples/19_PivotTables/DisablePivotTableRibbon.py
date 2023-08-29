from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "DisablePivotTableRibbon.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#Disable ribbon for this pivot table
pt.EnableWizard = False
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
