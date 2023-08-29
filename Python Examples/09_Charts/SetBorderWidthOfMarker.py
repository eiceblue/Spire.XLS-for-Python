from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetBorderWidthOfMarker.xlsx"
outputFile = "SetBorderWidthOfMarker.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
chart.Series[0].DataFormat.MarkerBorderWidth = 1.5 #unit is pt
chart.Series[1].DataFormat.MarkerBorderWidth = 2.5 #unit is pt
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

