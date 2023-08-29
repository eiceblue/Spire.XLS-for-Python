from spire.xls import *
from spire.common import *


outputFile = "CustomDataMarker.xlsx"

#Create a workbook
workbook = Workbook()
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
#Add some sample data
sheet.Name = "Demo"
sheet.Range["A1"].Value = "Tom"
sheet.Range["A2"].NumberValue = 1.5
sheet.Range["A3"].NumberValue = 2.1
sheet.Range["A4"].NumberValue = 3.6
sheet.Range["A5"].NumberValue = 5.2
sheet.Range["A6"].NumberValue = 7.3
sheet.Range["A7"].NumberValue = 3.1
sheet.Range["B1"].Value = "Kitty"
sheet.Range["B2"].NumberValue = 2.5
sheet.Range["B3"].NumberValue = 4.2
sheet.Range["B4"].NumberValue = 1.3
sheet.Range["B5"].NumberValue = 3.2
sheet.Range["B6"].NumberValue = 6.2
sheet.Range["B7"].NumberValue = 4.7
#Create a Scatter-Markers chart based on the sample data
chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers)
chart.DataRange = sheet.Range["A1:B7"]
chart.PlotArea.Visible = False
chart.SeriesDataFromRange = False
chart.TopRow = 5
chart.BottomRow = 22
chart.LeftColumn = 4
chart.RightColumn = 11
chart.ChartTitle = "Chart with Markers"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 10
#Format the markers in the chart by setting the background color, foreground color, type, size and transparency
cs1 = chart.Series[0]
cs1.DataFormat.MarkerBackgroundColor = Color.get_RoyalBlue()
cs1.DataFormat.MarkerForegroundColor = Color.get_WhiteSmoke()
cs1.DataFormat.MarkerSize = 7
cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign
cs1.DataFormat.MarkerTransparencyValue = 0.8
cs2 = chart.Series[1]
cs2.DataFormat.MarkerBackgroundColor = Color.get_Pink()
cs2.DataFormat.MarkerSize = 9
cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle
cs2.DataFormat.MarkerTransparencyValue = 0.9
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

