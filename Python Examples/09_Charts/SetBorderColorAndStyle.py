
from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample3.xlsx"
outputFile = "SetBorderColorAndStyle.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet from workbook and then get the first chart from the worksheet
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
#Set CustomLineWeight property for Series line
( chart.Series[0].DataPoints[0].DataFormat.LineProperties if isinstance(chart.Series[0].DataPoints[0].DataFormat.LineProperties, XlsChartBorder) else None).CustomLineWeight = 2.5
#Set color property for Series line
( chart.Series[0].DataPoints[0].DataFormat.LineProperties if isinstance(chart.Series[0].DataPoints[0].DataFormat.LineProperties, XlsChartBorder) else None).Color = Color.get_Red()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

