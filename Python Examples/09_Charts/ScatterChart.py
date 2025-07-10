from spire.xls import *
from spire.xls.common import *


def CreateChartData( sheet):
    #Set value of specified cell
    sheet.Range["A1"].Value = "Y(Salary)"
    sheet.Range["A2"].Value = "42763"
    sheet.Range["A3"].Value = "195387"
    sheet.Range["A4"].Value = "35672"
    sheet.Range["A5"].Value = "217637"
    sheet.Range["A6"].Value = "74734"
    sheet.Range["A7"].Value = "130550"
    sheet.Range["A8"].Value = "42976"
    sheet.Range["A9"].Value = "15132"
    sheet.Range["A10"].Value = "54936"
    sheet.Range["B1"].Value = "X(Car Price)"
    sheet.Range["B2"].Value = "19455"
    sheet.Range["B3"].Value = "93965"
    sheet.Range["B4"].Value = "20858"
    sheet.Range["B5"].Value = "107164"
    sheet.Range["B6"].Value = "34036"
    sheet.Range["B7"].Value = "87806"
    sheet.Range["B8"].Value = "17927"
    sheet.Range["B9"].Value = "61518"
    sheet.Range["B10"].Value = "29479"
    #Style
    sheet.Range["A1:B1"].ColumnWidth = 12
    sheet.Range["A1:B1"].RowHeight = 15
    sheet.Range["A1:B1"].Style.Color = Color.get_DarkGray()
    sheet.Range["A1:B1"].Style.Font.Color = Color.get_White()
    sheet.Range["A1:B1"].Style.VerticalAlignment = VerticalAlignType.Center
    sheet.Range["A1:B1"].Style.HorizontalAlignment = HorizontalAlignType.Center
    sheet.Range["A2:B10"].Style.NumberFormat = "\"$\"#,##0"

outputFile = "ScatterChart.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "Scatter Chart"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B10"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 11
chart.RightColumn = 10
chart.BottomRow = 28
chart.ChartTitle = "Scatter Chart"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.Series[0].CategoryLabels = sheet.Range["A2:A10"]
chart.Series[0].Values = sheet.Range["B2:B10"]
#Add a trend line for the first series
chart.Series[0].TrendLines.Add(TrendLineType.Exponential)
chart.PrimaryValueAxis.Title = "Salary"
chart.PrimaryCategoryAxis.Title = "Car Price"
workbook.SaveToFile(outputFile, FileFormat.Version2010)
workbook.Dispose()
