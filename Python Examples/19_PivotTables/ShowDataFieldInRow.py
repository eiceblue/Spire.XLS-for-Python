from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "ShowDataFieldInRow.xlsx"

#create a workbook
workbook = Workbook()
#load an excel document including pivot table
workbook.LoadFromFile(inputFile)
sheet=workbook.Worksheets[1]
#get the data in Pivot Table
pivotTable = sheet.PivotTables[0] if isinstance(sheet.PivotTables[0], XlsPivotTable) else None
#show the datafield in row
pivotTable.ShowDataFieldInRow = True
#calculate data
pivotTable.CalculateData()

#save the file
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()


