from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Funnel.xlsx"
outputFile = "CreateFunnelChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Funnel
officeChart.ChartType = ExcelChartType.Funnel

#set data range in the worksheet
officeChart.DataRange = sheet.Range["A1:B6"]

#set the chart title
officeChart.ChartTitle = "Funnel"

#format the legend and data label option
officeChart.HasLegend = False
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = True
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()


