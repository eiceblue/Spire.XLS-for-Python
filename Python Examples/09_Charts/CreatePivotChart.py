from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PivotTable.xlsx"
outputFile = "CreatePivotChart.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#get the first worksheet
sheet = workbook.Worksheets[0]
#get the first pivot table in the worksheet
pivotTable = sheet.PivotTables[0]
#create a clustered column chart based on the pivot table
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable)
#set chart position
chart.TopRow = 10
chart.LeftColumn = 1
chart.RightColumn = 7
chart.BottomRow = 25
#set chart title
chart.ChartTitle = "Pivot Chart"
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

