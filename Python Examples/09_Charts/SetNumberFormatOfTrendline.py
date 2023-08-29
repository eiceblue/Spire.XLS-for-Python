from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample4.xlsx"
outputFile = "SetNumberFormatOfTrendline.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
#Get the trendline of the chart and then extract the equation of the trendline
trendLine = chart.Series[1].TrendLines[0]
#Set the number format of trendLine to "#,##0.00"
trendLine.DataLabel.NumberFormat = "#,##0.00"
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
