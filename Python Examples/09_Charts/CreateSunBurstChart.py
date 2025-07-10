from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SunBurst.xlsx"
outputFile = "CreateSunBurstChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as SunBurst
officeChart.ChartType = ExcelChartType.SunBurst

#set data range in the worksheet   
officeChart.DataRange = sheet["A1:D16"]

officeChart.TopRow = 1
officeChart.BottomRow = 17
officeChart.LeftColumn = 6
officeChart.RightColumn = 14

#set the chart title
officeChart.ChartTitle = "Sales by quarter"

#format data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8

#hide the legend
officeChart.HasLegend = False

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()