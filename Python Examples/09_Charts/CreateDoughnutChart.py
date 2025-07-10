from spire.xls import *
from spire.xls.common import *


outputFile = "CreateDoughnutChart.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Insert data
sheet.Range["A1"].Value = "Country"
sheet.Range["A1"].Style.Font.IsBold = True
sheet.Range["A2"].Value = "Cuba"
sheet.Range["A3"].Value = "Mexico"
sheet.Range["A4"].Value = "France"
sheet.Range["A5"].Value = "German"
sheet.Range["B1"].Value = "Sales"
sheet.Range["B1"].Style.Font.IsBold = True
sheet.Range["B2"].NumberValue = 6000
sheet.Range["B3"].NumberValue = 8000
sheet.Range["B4"].NumberValue = 9000
sheet.Range["B5"].NumberValue = 8500
#Add a new chart, set chart type as doughnut
chart = sheet.Charts.Add()
chart.ChartType = ExcelChartType.Doughnut
chart.DataRange = sheet.Range["A1:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 4
chart.TopRow = 2
chart.RightColumn = 12
chart.BottomRow = 22
#Chart title
chart.ChartTitle = "Market share by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = True
chart.Legend.Position = LegendPositionType.Top
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

