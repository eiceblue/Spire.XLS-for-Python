from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "SetColorForChartArea.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Set color for chart area
chart.ChartArea.Fill.ForeColor = Color.get_LightSeaGreen()
#Set color for plot area
chart.PlotArea.Fill.ForeColor = Color.get_LightGray()
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

