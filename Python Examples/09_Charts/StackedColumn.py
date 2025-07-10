from spire.xls import *
from spire.xls.common import *


def CreateChartData( sheet):
    #Set value of specified cell
    sheet.Range["A1"].Value = "Country"
    sheet.Range["A2"].Value = "Cuba"
    sheet.Range["A3"].Value = "Mexico"
    sheet.Range["A4"].Value = "France"
    sheet.Range["A5"].Value = "German"
    sheet.Range["B1"].Value = "Jun"
    sheet.Range["B2"].NumberValue = 6000
    sheet.Range["B3"].NumberValue = 8000
    sheet.Range["B4"].NumberValue = 9000
    sheet.Range["B5"].NumberValue = 8500
    sheet.Range["C1"].Value = "Aug"
    sheet.Range["C2"].NumberValue = 3000
    sheet.Range["C3"].NumberValue = 2000
    sheet.Range["C4"].NumberValue = 2300
    sheet.Range["C5"].NumberValue = 4200
    #Style
    sheet.Range["A1:C1"].RowHeight = 15
    sheet.Range["A1:C1"].Style.Color = Color.get_DarkGray()
    sheet.Range["A1:C1"].Style.Font.Color = Color.get_White()
    sheet.Range["A1:C1"].Style.VerticalAlignment = VerticalAlignType.Center
    sheet.Range["A1:C1"].Style.HorizontalAlignment = HorizontalAlignType.Center
    sheet.Range["B2:C5"].Style.NumberFormat = "\"$\"#,##0"

outputFile1 = "StackedColumn.xlsx"
outputFile2 = "StackedColumn_3D.xlsx"

#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "StackedColumn"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.ColumnStacked
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axes
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
#Save and Launch
workbook.SaveToFile(outputFile1, ExcelVersion.Version2010)
workbook.Dispose()

#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "StackedColumn"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Column3DStacked
#Chart title
chart.ChartTitle = "Sales market by country"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Chart Axes
chart.PrimaryCategoryAxis.Title = "Country"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
for cs in chart.Series:
    cs.Format.Options.IsVaryColor = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
#Chart Legend
chart.Legend.Position = LegendPositionType.Top
#Save and Launch
workbook.SaveToFile(outputFile2, ExcelVersion.Version2010)
workbook.Dispose()
