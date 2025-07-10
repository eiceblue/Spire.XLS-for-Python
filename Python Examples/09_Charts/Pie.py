from spire.xls import *
from spire.xls.common import *

def CreateChartData( sheet):
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


outputFile1 = "Pie.xlsx"
outputFile2 =  "Pie_3D.xlsx"

#Pie
#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "Pie Chart"
#Add a chart
chart = None
chart = sheet.Charts.Add(ExcelChartType.Pie)
#Set chart data
CreateChartData(sheet)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 9
chart.BottomRow = 25
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Values = sheet.Range["B2:B5"]
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
#Save and Launch
workbook.SaveToFile(outputFile1, ExcelVersion.Version2010)
workbook.Dispose()


#Pie_3D
#Create a Workbook
workbook = Workbook()
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
sheet.Name = "Pie Chart"
#Add a chart
chart = None
chart = sheet.Charts.Add(ExcelChartType.Pie3D)
#Set chart data
CreateChartData(sheet)
#Set region of chart data
chart.DataRange = sheet.Range["B2:B5"]
chart.SeriesDataFromRange = False
#Set position of chart
chart.LeftColumn = 1
chart.TopRow = 6
chart.RightColumn = 9
chart.BottomRow = 25
#Chart title
chart.ChartTitle = "Sales by year"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs = chart.Series[0]
cs.CategoryLabels = sheet.Range["A2:A5"]
cs.Values = sheet.Range["B2:B5"]
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
chart.PlotArea.Fill.Visible = False
#Save and Launch
workbook.SaveToFile(outputFile2, ExcelVersion.Version2010)
workbook.Dispose()

