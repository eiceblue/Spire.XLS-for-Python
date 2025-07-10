from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "DeleteLegend.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Delete the first and the second legend entries from the chart
chart.Legend.LegendEntries[0].Delete()
chart.Legend.LegendEntries[1].Delete()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

