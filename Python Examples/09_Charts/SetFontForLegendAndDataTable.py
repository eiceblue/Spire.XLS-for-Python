from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "SetFontForLegendAndDataTable.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet from workbook
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
#Create a font with specified size and color
font = workbook.CreateFont()
font.Size = 14.0
font.Color = Color.get_Red()
#Apply the font to chart Legend
chart.Legend.TextArea.SetFont(font)
#Apply the font to chart DataLabel
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

