from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "PivotTableExample.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Modify data of data source
data = workbook.Worksheets["Data"]
data.Range["A2"].Text = "NewValue"
data.Range["D2"].NumberValue = 28000
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Refresh and calculate
pt.Cache.IsRefreshOnLoad = True
pt.CalculateData()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

