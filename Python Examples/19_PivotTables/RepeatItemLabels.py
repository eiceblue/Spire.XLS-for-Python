from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/RepeatItemLabelsExample.xlsx"
outputFile = "RepeatItemLabels.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add an empty worksheet 
sheet2 = workbook.CreateEmptySheet()
#Add PivotTable
sheet2.Name = "Pivot Table"
dataRange = sheet.Range["A1:D9"]
cache = workbook.PivotCaches.Add(dataRange)
pt = sheet2.PivotTables.Add("Pivot Table", sheet.Range["A1"], cache)
r1 = pt.PivotFields["VendorNo"]
r1.Axis = AxisTypes.Row
pt.Options.RowHeaderCaption = "VendorNo"
r1.Subtotals = SubtotalTypes.none
r1.RepeatItemLabels = True
#Repeat item lables
pt.PivotFields["OnHand"].RepeatItemLabels = True
pt.Options.RowLayout = PivotTableLayoutType.Tabular
r2 = pt.PivotFields["Desc"]
r2.Axis = AxisTypes.Row
pt.DataFields.Add(pt.PivotFields["OnHand"], "Sum of onHand", SubtotalTypes.none)
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium12
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

