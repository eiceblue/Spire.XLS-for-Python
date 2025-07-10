from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
inputImg ="./Demos/Data/background.png"
outputFile = "FillChartElementWithPicture_A.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet from workbook
ws = workbook.Worksheets[0]
#Get the first chart
chart = ws.Charts[0]
# A. Fill chart area with image
chart.ChartArea.Fill.CustomPicture(Stream(inputImg), "None")
chart.PlotArea.Fill.Transparency = 0.9
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

