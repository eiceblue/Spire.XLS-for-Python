from spire.xls import *
from spire.common import *


outputFile = "AddErrorBars.xlsx"

#Create a workbook
workbook = Workbook()
#Create a empty sheet
workbook.CreateEmptySheets(1)
#Add data
sheet = workbook.Worksheets[0]
sheet.Name = "Demo"
sheet.Range["A1"].Value = "Month"
sheet.Range["A2"].Value = "Jan."
sheet.Range["A3"].Value = "Feb."
sheet.Range["A4"].Value = "Mar."
sheet.Range["A5"].Value = "Apr."
sheet.Range["A6"].Value = "May."
sheet.Range["A7"].Value = "Jun."
sheet.Range["B1"].Value = "Planned"
sheet.Range["B2"].NumberValue = 3.3
sheet.Range["B3"].NumberValue = 2.5
sheet.Range["B4"].NumberValue = 2.0
sheet.Range["B5"].NumberValue = 3.7
sheet.Range["B6"].NumberValue = 4.5
sheet.Range["B7"].NumberValue = 4.0
sheet.Range["C1"].Value = "Actual"
sheet.Range["C2"].NumberValue = 3.8
sheet.Range["C3"].NumberValue = 3.2
sheet.Range["C4"].NumberValue = 1.7
sheet.Range["C5"].NumberValue = 3.5
sheet.Range["C6"].NumberValue = 4.5
sheet.Range["C7"].NumberValue = 4.3
#Add a line chart and then add percentage error bar to the chart
chart = sheet.Charts.Add(ExcelChartType.Line)
chart.DataRange = sheet.Range["B1:B7"]
chart.SeriesDataFromRange = False
#Set chart position
chart.TopRow = 8
chart.BottomRow = 25
chart.LeftColumn = 2
chart.RightColumn = 9
chart.ChartTitle = "Error Bar 10% Plus"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
cs1 = chart.Series[0]
cs1.CategoryLabels = sheet.Range["A2:A7"]
cs1.ErrorBar(True, ErrorBarIncludeType.Plus, ErrorBarType.Percentage, 10.0)
# Add a column chart with standard error bars as comparison
chart2 = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart2.DataRange = sheet.Range["B1:C7"]
chart2.SeriesDataFromRange = False
#Set chart position
chart2.TopRow = 8
chart2.BottomRow = 25
chart2.LeftColumn = 10
chart2.RightColumn = 17
chart2.ChartTitle = "Standard Error Bar"
chart2.ChartTitleArea.IsBold = True
chart2.ChartTitleArea.Size = 12
cs2 = chart2.Series[0]
cs2.CategoryLabels = sheet.Range["A2:A7"]
cs2.ErrorBar(True, ErrorBarIncludeType.Minus, ErrorBarType.StandardError, 0.3)
cs3 = chart2.Series[1]
cs3.ErrorBar(True, ErrorBarIncludeType.Both, ErrorBarType.StandardError, 0.5)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

