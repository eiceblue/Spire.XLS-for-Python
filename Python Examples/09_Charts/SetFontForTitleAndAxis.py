from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "SetFontForTitleAndAxis.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Set font for chart title and chart axis
worksheet = workbook.Worksheets[0]
chart = worksheet.Charts[0]
#Format the font for the chart title
chart.ChartTitleArea.Color = Color.get_Blue()
chart.ChartTitleArea.Size = 20.0
#Format the font for the chart Axis
chart.PrimaryValueAxis.Font.Color = Color.get_Gold()
chart.PrimaryValueAxis.Font.Size = 10.0
chart.PrimaryCategoryAxis.Font.Color = Color.get_Red()
chart.PrimaryCategoryAxis.Font.Size = 20.0
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

