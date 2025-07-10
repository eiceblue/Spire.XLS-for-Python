from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/TreeMap.xlsx"
outputFile = "CreateTreeMapChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set chart type as TreeMap
officeChart.ChartType = ExcelChartType.TreeMap

#set data range in the worksheet   
officeChart.DataRange = sheet["A2:C11"]
officeChart.TopRow = 1
officeChart.BottomRow = 19
officeChart.LeftColumn = 4
officeChart.RightColumn = 14

#Set the chart title
officeChart.ChartTitle = "Area by countries"

#set the Treemap label option
officeChart.Series[0].DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner

#format data labels      
officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()