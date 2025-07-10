from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "SetFormatOptions.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Set the PivotTable report is automatically formatted
pt.Options.IsAutoFormat = True
#Setting the PivotTable report shows grand totals for rows.
pt.ShowRowGrand = True
#Setting the PivotTable report shows grand totals for columns.
pt.ShowColumnGrand = True
#Setting the PivotTable report displays a custom string in cells that contain null values.
pt.DisplayNullString = True
pt.NullString = "null"
#Setting the PivotTable report's layout
pt.PageFieldOrder = PagesOrderType.DownThenOver
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

