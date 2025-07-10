from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/HistogramChart.xlsx"
outputFile = "CreateHistogramChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as Histogram
officeChart.ChartType = ExcelChartType.Histogram

#set data range in the worksheet   
officeChart.DataRange = sheet["A1:A15"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 12

#category axis bin settings        
officeChart.PrimaryCategoryAxis.BinWidth = 8

#gap width settings
officeChart.Series[0].DataFormat.Options.GapWidth = 6

#set the chart title and axis title
officeChart.ChartTitle = "Height Data"
officeChart.PrimaryValueAxis.Title = "Number of students"
officeChart.PrimaryCategoryAxis.Title = "Height"

#hide the legend
officeChart.HasLegend = False

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()
