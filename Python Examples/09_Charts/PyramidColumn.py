from spire.xls import *
from spire.common import *


def CreateChartData(sheet):
    #Set value of specified cell
    sheet.Range["A1"].Value = "Year"
    sheet.Range["A2"].Value = "2002"
    sheet.Range["A3"].Value = "2003"
    sheet.Range["A4"].Value = "2004"
    sheet.Range["A5"].Value = "2005"
    sheet.Range["B1"].Value = "Sales"
    sheet.Range["B2"].NumberValue = 4000
    sheet.Range["B3"].NumberValue = 6000
    sheet.Range["B4"].NumberValue = 7000
    sheet.Range["B5"].NumberValue = 8500
    #Style
    sheet.Range["A1:B1"].RowHeight = 15
    sheet.Range["A1:B1"].Style.Color = Color.get_DarkGray()
    sheet.Range["A1:B1"].Style.Font.Color = Color.get_White()
    sheet.Range["A1:B1"].Style.VerticalAlignment = VerticalAlignType.Center
    sheet.Range["A1:B1"].Style.HorizontalAlignment = HorizontalAlignType.Center
    sheet.Range["B2:C5"].Style.NumberFormat = "\"$\"#,##0"


outputFile1 =  "PyramidColumn.xlsx"
outputFile2 =  "PyramidColumn_3D.xlsx"


#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "Chart"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Pyramid3DClustered
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Year"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Format.Options.IsVaryColor = True
chart.Legend.Position = LegendPositionType.Top
workbook.SaveToFile(outputFile1, ExcelVersion.Version2010)
workbook.Dispose()


#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "Chart"
#Set chart data
CreateChartData(sheet)
#Add a chart
chart = sheet.Charts.Add()
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 11
chart.BottomRow = 29
chart.ChartType = ExcelChartType.Pyramid3DClustered
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
chart.PrimaryCategoryAxis.Title = "Year"
chart.PrimaryCategoryAxis.Font.IsBold = True
chart.PrimaryCategoryAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
chart.PrimaryValueAxis.HasMajorGridLines = False
chart.PrimaryValueAxis.MinValue = 1000
chart.PrimaryValueAxis.TitleArea.IsBold = True
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Format.Options.IsVaryColor = True
chart.Legend.Position = LegendPositionType.Top
workbook.SaveToFile(outputFile2, ExcelVersion.Version2010)
workbook.Dispose()

