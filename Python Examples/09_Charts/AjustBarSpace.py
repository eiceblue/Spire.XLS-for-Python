from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "AjustBarSpace.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet from workbook and then get the first chart from the worksheet
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
#Ajust the space between bars
for cs in chart.Series:
    cs.Format.Options.GapWidth = 200
    cs.Format.Options.Overlap = 0
#Save the document        
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

