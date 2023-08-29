from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SortPivotTable.xlsx"
outputFile = "SortPivotTable.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add an empty worksheet 
sheet2 = workbook.CreateEmptySheet()
sheet2.Name = "Pivot Table"
#Specify the datasorce
dataRange = sheet.Range["A1:C9"]
cache = workbook.PivotCaches.Add(dataRange)
#Add PivotTable
pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache)
r1 = pt.PivotFields["No"]
r1.Axis = AxisTypes.Row
pt.Options.RowLayout = PivotTableLayoutType.Tabular
#Sort PivotField
r1.SortType = PivotFieldSortType.Descending
r2 = pt.PivotFields["Name"]
r2.Axis = AxisTypes.Row
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.none)
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


