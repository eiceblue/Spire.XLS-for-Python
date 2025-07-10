from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WaterfallChart.xlsx"
outputFile = "CreateWaterfallChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as WaterFall
officeChart.ChartType = ExcelChartType.WaterFall

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:B8"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12

#set data point as total in chart
officeChart.Series[0].DataPoints[3].SetAsTotal = True
officeChart.Series[0].DataPoints[6].SetAsTotal = True

#show the connector lines between data points
officeChart.Series[0].Format.ShowConnectorLines = True

#set the chart title
officeChart.ChartTitle = "WaterFall Chart"

#format data label and legend option
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = True
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8
officeChart.Legend.Position = LegendPositionType.Right

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()