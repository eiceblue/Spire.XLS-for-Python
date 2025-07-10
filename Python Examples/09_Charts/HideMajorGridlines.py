from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "HideMajorGridlines.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Hide major gridlines
chart.PrimaryValueAxis.HasMajorGridLines = False
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

