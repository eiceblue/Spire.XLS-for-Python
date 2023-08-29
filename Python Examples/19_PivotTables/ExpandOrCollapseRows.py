from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_7.xlsx"
outputFile = "ExpandOrCollapseRows.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the data in Pivot Table.
pivotTable = sheet.PivotTables[0]
#Calculate Data.
pivotTable.CalculateData()
#Collapse the rows.
(pivotTable.PivotFields["Vendor No"]).HideItemDetail("3501", True)
#Expand the rows.
( pivotTable.PivotFields["Vendor No"]).HideItemDetail("3502", False)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

