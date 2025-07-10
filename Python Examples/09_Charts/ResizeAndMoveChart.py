from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "ResizeAndMoveChart.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chart from the first worksheet
sheet = workbook.Worksheets[0]
chart = sheet.Charts[0]
#Set position of the chart
chart.LeftColumn = 5
chart.TopRow = 1
#Resize the chart
chart.Width = 500
chart.Height = 350
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
