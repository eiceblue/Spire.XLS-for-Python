from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CreateBubbleChart.xlsx"
outputFile = "CreateBubbleChart.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet and set its name
sheet = workbook.Worksheets[0]
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.Bubble)
#Set region of chart data
chart.DataRange = sheet.Range["A1:C5"]
chart.SeriesDataFromRange = False
chart.Series[0].Bubbles = sheet.Range["C2:C5"]
#Set position of chart
chart.LeftColumn = 7
chart.TopRow = 6
chart.RightColumn = 16
chart.BottomRow = 29
chart.ChartTitle = "Bubble Chart"
chart.ChartTitleArea.IsBold = True
chart.ChartTitleArea.Size = 12
#Save the Excel file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

