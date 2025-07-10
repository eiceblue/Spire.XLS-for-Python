from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/PivotTableExample.xlsx"
outputFile = "FormattingAppearance.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file including pivot table
workbook.LoadFromFile(inputFile)
#Get the sheet in which the pivot table is located
sheet = workbook.Worksheets["PivotTable"]
pt = sheet.PivotTables[0]
#Format appearance
pt.BuiltInStyle = PivotBuiltInStyles.PivotStyleLight10
pt.Options.ShowGridDropZone = True
pt.Options.RowLayout = PivotTableLayoutType.Tabular
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

