from spire.xls import *
from spire.common import *


def CreateChartData( sheet):
    #Set value of specified cell
    sheet.Range["A1"].Value = "Month"
    sheet.Range["A2"].Value = "Jan"
    sheet.Range["A3"].Value = "Feb"
    sheet.Range["A4"].Value = "Mar"
    sheet.Range["A5"].Value = "Apr"
    sheet.Range["A6"].Value = "May"
    sheet.Range["A7"].Value = "Jun"
    sheet.Range["A8"].Value = "Jul"
    sheet.Range["A9"].Value = "Aug"
    sheet.Range["B1"].Value = "Planned"
    sheet.Range["B2"].NumberValue = 38
    sheet.Range["B3"].NumberValue = 47
    sheet.Range["B4"].NumberValue = 39
    sheet.Range["B5"].NumberValue = 36
    sheet.Range["B6"].NumberValue = 27
    sheet.Range["B7"].NumberValue = 25
    sheet.Range["B8"].NumberValue = 36
    sheet.Range["B9"].NumberValue = 48


outputFile = "FormatAxis.xlsx"

#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "FormatAxis"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart.DataRange = sheet.Range["B1:B9"]
chart.SeriesDataFromRange = False
chart.PlotArea.Visible = False
chart.TopRow = 10
chart.BottomRow = 28
chart.LeftColumn = 2
chart.RightColumn = 10
chart.ChartTitle = "Chart with Customized Axis"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A9"]
#Format axis
chart.PrimaryValueAxis.MajorUnit = 8
chart.PrimaryValueAxis.MinorUnit = 2
chart.PrimaryValueAxis.MaxValue = 50
chart.PrimaryValueAxis.MinValue = 0
chart.PrimaryValueAxis.IsReverseOrder = False
chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside
chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside
chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis
chart.PrimaryValueAxis.CrossesAt = 0
#Set NumberFormat
chart.PrimaryValueAxis.NumberFormat = "$#,##0"
chart.PrimaryValueAxis.IsSourceLinked = False
serie = chart.Series[0]
p = serie.DataPoints[0]
p = serie.DataPoints[1]
#p = serie.DataPoints[2]
#p = serie.DataPoints[3]
for dataPoint in serie.DataPoints:
    #Format Series
    dataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor
    dataPoint.DataFormat.Fill.ForeColor = Color.get_LightGreen()
for dataPoint in serie.DataPoints:
    #Set transparency
    dataPoint.DataFormat.Fill.Transparency = 0.3
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

