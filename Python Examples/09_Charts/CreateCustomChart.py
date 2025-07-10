from spire.xls import *
from spire.xls.common import *

outputFile = "CreateCustomChart.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Set values
sheet.Range["A1"].Value = "60"
sheet.Range["A2"].Value = "90"
sheet.Range["A3"].Value = "80"
sheet.Range["A4"].Value = "85"
sheet.Range["B1"].Value = "100"
sheet.Range["B2"].Value = "110"
sheet.Range["B3"].Value = "80"
sheet.Range["B4"].Value = "70"
#Add a chart based on the data from A1 to B4
chart = sheet.Charts.Add()
chart.DataRange = sheet.Range["A1:B4"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 10
chart.RightColumn = 7
chart.BottomRow = 25
#Apply different chart type to different series
cs1 = chart.Series[0]
cs1.SerieType = ExcelChartType.ColumnClustered
cs2 = chart.Series[1]
cs2.SerieType = ExcelChartType.Line
chart.ChartTitle = "Custom chart"
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

