from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SetFont.xlsx"
outputFile = "SetFont.xlsx"

#Load a Workbook from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first sheet
chart = sheet.Charts[0]
#Create a font
font = workbook.CreateFont()
font.Size = 15.0
font.Color = Color.get_LightSeaGreen()
for cs in chart.Series:
    #Set font
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

