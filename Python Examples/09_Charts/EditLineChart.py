from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/LineChart.xlsx"
outputFile = "LineChart.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the line chart
chart = sheet.Charts[0]
#Add a new series
cs = chart.Series.Add("Added")
#Set the values for the series
cs.Values = sheet.Range["I1:L1"]
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


