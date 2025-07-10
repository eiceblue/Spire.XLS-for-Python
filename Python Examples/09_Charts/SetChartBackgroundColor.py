from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "SetChartBackgroundColor.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet from workbook and then get the first chart from the worksheet
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
#Set background color
chart.ChartArea.ForeGroundColor = Color.get_LightYellow()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

