from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/DataCallout.xlsx"
outputFile = "DataCallout.xlsx"

#Create a Workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
    cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = True
#Save and Launch
workbook.SaveToFile(outputFile, FileFormat.Version2010)
workbook.Dispose()

