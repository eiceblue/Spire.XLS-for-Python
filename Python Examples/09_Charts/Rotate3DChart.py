from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample3.xlsx"
outputFile = "Rotate3DChart.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chart from the first worksheet
sheet = workbook.Worksheets[0]
chart = sheet.Charts[0]
#X rotation:
chart.Rotation = 30
#Y rotation:
chart.Elevation = 20
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

