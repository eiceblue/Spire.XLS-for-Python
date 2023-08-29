from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample3.xlsx"
outputFile = "SoftEdgesEffect.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Specify the size of the soft edge. Value can be set from 0 to 100
chart.ChartArea.Shadow.SoftEdge = 25
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


