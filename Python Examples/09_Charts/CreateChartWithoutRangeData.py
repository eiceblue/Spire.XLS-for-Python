from spire.xls import *
from spire.xls.common import *

outputFile = "CreateChartWithoutRangeData.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Add a chart to the worksheet
chart = sheet.Charts.Add()
chart.ChartTitle = "Sample Chart"
#Add a series to the chart
series = chart.Series.Add()
#Add data 
series.EnteredDirectlyValues = [Int32(10), Int32(20), Int32(30)]
v = series.EnteredDirectlyValues
#Save the document      
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
