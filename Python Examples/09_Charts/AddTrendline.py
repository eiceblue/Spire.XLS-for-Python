from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample2.xlsx"
outputFile = "AddTrendline.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#select chart and set logarithmic trendline
chart = sheet.Charts[0]
chart.ChartTitle = "Logarithmic Trendline"
chart.Series[0].TrendLines.Add(TrendLineType.Logarithmic)
#select chart and set moving_average trendline
chart1 = sheet.Charts[1]
chart1.ChartTitle = "Moving Average Trendline"
chart1.Series[0].TrendLines.Add(TrendLineType.Moving_Average)
#select chart and set linear trendline
chart2 = sheet.Charts[2]
chart2.ChartTitle = "Linear Trendline"
chart2.Series[0].TrendLines.Add(TrendLineType.Linear)
#select chart and set exponential trendline
chart3 = sheet.Charts[3]
chart3.ChartTitle = "Exponential Trendline"
chart3.Series[0].TrendLines.Add(TrendLineType.Exponential)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

