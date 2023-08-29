from spire.xls import *
from spire.common import *


def CreateChartData(sheet):
    #Product
    sheet.Range["A1"].Value = "Product"
    sheet.Range["A2"].Value = "Bikes"
    sheet.Range["A3"].Value = "Cars"
    sheet.Range["A4"].Value = "Trucks"
    sheet.Range["A5"].Value = "Buses"
    #Paris
    sheet.Range["B1"].Value = "Paris"
    sheet.Range["B2"].NumberValue = 4000
    sheet.Range["B3"].NumberValue = 23000
    sheet.Range["B4"].NumberValue = 4000
    sheet.Range["B5"].NumberValue = 30000
    #New York
    sheet.Range["C1"].Value = "New York"
    sheet.Range["C2"].NumberValue = 30000
    sheet.Range["C3"].NumberValue = 7600
    sheet.Range["C4"].NumberValue = 18000
    sheet.Range["C5"].NumberValue = 8000
    #Style
    sheet.Range["A1:C1"].Style.Font.IsBold = True
    sheet.Range["A2:C2"].Style.KnownColor = ExcelColors.LightYellow
    sheet.Range["A3:C3"].Style.KnownColor = ExcelColors.LightGreen1
    sheet.Range["A4:C4"].Style.KnownColor = ExcelColors.LightOrange
    sheet.Range["A5:C5"].Style.KnownColor = ExcelColors.LightTurquoise
    #Border
    style = sheet.Range["A1:C5"].Style
    borders = style.Borders
    topborder = borders[BordersLineType.EdgeTop]
    topborder.Color = Color.FromRgb(0, 0, 128)
    borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin
    borders[BordersLineType.EdgeBottom].Color = Color.FromRgb(0, 0, 128)
    sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin
    sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeLeft].Color = Color.FromRgb(0, 0, 128)
    sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin
    sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeRight].Color = Color.FromRgb(0, 0, 128)
    sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin
    sheet.Range["B2:C5"].Style.NumberFormat = "\"$\"#,##0"

outputFile1 = "CreateRadarChart.xlsx"
outputFile2 =  "CreateRadarChart_Fill.xlsx"

#Radar
#Create a workbook
workbook = Workbook()
#Initailize worksheet
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
sheet.Name = "Chart data"
sheet.GridLinesVisible = False
#Writes chart data
CreateChartData(sheet)
#Add a new  chart worsheet to workbook
chart = sheet.Charts.Add()
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.ChartType = ExcelChartType.Radar
#Chart title
chart.ChartTitle = "Sale market by region"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Corner
#Save the document
workbook.SaveToFile(outputFile1, ExcelVersion.Version2013)
workbook.Dispose()

#RadarFilled
#Create a workbook
workbook = Workbook()
#Initailize worksheet
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
sheet.Name = "Chart data"
sheet.GridLinesVisible = False
#Writes chart data
CreateChartData(sheet)
#Add a new  chart worsheet to workbook
chart = sheet.Charts.Add()
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.ChartType = ExcelChartType.RadarFilled
#Chart title
chart.ChartTitle = "Sale market by region"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PlotArea.Fill.Visible = False
chart.Legend.Position = LegendPositionType.Corner
#Save the document
workbook.SaveToFile(outputFile2, ExcelVersion.Version2013)
workbook.Dispose()
