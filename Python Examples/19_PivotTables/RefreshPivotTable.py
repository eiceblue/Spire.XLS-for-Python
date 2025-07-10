from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_7.xlsx"
outputFile = "RefreshPivotTable.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[1]
#Update the data source of PivotTable.
sheet.Range["D2"].Value = "999"
#Get the PivotTable that was built on the data source.
pt = workbook.Worksheets[0].PivotTables[0]
#Refresh the data of PivotTable.
pt.Cache.IsRefreshOnLoad = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

