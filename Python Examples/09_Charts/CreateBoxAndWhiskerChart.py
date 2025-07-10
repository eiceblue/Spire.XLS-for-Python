from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/BoxAndWhiskerChart.xlsx"
outputFile = "CreateBoxAndWhiskerChart.xlsx"

#create a workbook
workbook = Workbook()

#load an excel document
workbook.LoadFromFile(inputFile)

#get the first worksheet
sheet = workbook.Worksheets[0]

#add a new chart
officeChart = sheet.Charts.Add()

#set the chart title
officeChart.ChartTitle = "Yearly Vehicle Sales"

#set chart type as Box and Whisker
officeChart.ChartType = ExcelChartType.BoxAndWhisker

#set data range in the worksheet
officeChart.DataRange = sheet["A1:E17"]

#box and Whisker settings on first series
seriesA = officeChart.Series[0]
seriesA.DataFormat.ShowInnerPoints = False
seriesA.DataFormat.ShowOutlierPoints = True
seriesA.DataFormat.ShowMeanMarkers = True
seriesA.DataFormat.ShowMeanLine = False
seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian

#box and Whisker settings on second series   
seriesB = officeChart.Series[1]
seriesB.DataFormat.ShowInnerPoints = False
seriesB.DataFormat.ShowOutlierPoints = True
seriesB.DataFormat.ShowMeanMarkers = True
seriesB.DataFormat.ShowMeanLine = False
seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian

#box and Whisker settings on third series   
seriesC = officeChart.Series[2]
seriesC.DataFormat.ShowInnerPoints = False
seriesC.DataFormat.ShowOutlierPoints = True
seriesC.DataFormat.ShowMeanMarkers = True
seriesC.DataFormat.ShowMeanLine = False
seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian

#save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
workbook.Dispose()

