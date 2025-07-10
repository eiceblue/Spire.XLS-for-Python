from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChangeDataSource.xlsx"
outputFile = "ChangeDataSource.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
Range = sheet.Range["A1:C15"]
table = workbook.Worksheets[1].PivotTables[0]
#Change data source
table.ChangeDataSource(Range)
table.Cache.IsRefreshOnLoad = False
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


