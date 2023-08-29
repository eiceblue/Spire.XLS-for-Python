from spire.xls import *
from spire.common import *

outputFile = "SetAndFormatDataLabel.xlsx"

#Create a workbook
workbook = Workbook()
workbook.CreateEmptySheets(1)
sheet = workbook.Worksheets[0]
sheet.Name = "Demo"
sheet.Range["A1"].Value = "Month"
sheet.Range["A2"].Value = "Jan"
sheet.Range["A3"].Value = "Feb"
sheet.Range["A4"].Value = "Mar"
sheet.Range["A5"].Value = "Apr"
sheet.Range["A6"].Value = "May"
sheet.Range["A7"].Value = "Jun"
sheet.Range["B1"].Value = "Peter"
sheet.Range["B2"].NumberValue = 25
sheet.Range["B3"].NumberValue = 18
sheet.Range["B4"].NumberValue = 8
sheet.Range["B5"].NumberValue = 13
sheet.Range["B6"].NumberValue = 22
sheet.Range["B7"].NumberValue = 28
chart = sheet.Charts.Add(ExcelChartType.LineMarkers)
chart.DataRange = sheet.Range["B1:B7"]
chart.PlotArea.Visible = False
chart.SeriesDataFromRange = False
chart.TopRow = 5
chart.BottomRow = 26
chart.LeftColumn = 2
chart.RightColumn = 11
chart.ChartTitle = "Data Labels Demo"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A7"]
cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = False
cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = False
cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". "
cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9
cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.get_Red()
cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri"
cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

