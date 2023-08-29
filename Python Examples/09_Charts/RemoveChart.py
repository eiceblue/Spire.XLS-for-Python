from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "RemoveChart.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet from the workbook
sheet = workbook.Worksheets[0]
#Get the first chart from the first worksheet
chart = sheet.Charts[0]
#Remove the chart
chart.Remove()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

