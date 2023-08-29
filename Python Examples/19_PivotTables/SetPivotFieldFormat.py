from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "SetPivotFieldFormat.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
pf = pt.PivotFields[0]
#Setting the field auto sort ascend.
pf.SortType = PivotFieldSortType.Ascending
#Setting Subtotal auto show.
pf.SubtotalTop = True
#Setting Subtotal as Count type
pf.Subtotals = SubtotalTypes.Count
#Setting the field auto show.
pf.IsAutoShow = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

