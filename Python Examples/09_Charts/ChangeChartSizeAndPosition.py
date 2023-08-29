from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "ChangeChartSizeAndPosition.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Change chart size
chart.Width = 600
chart.Height = 500
#Change chart position
chart.LeftColumn = 3
chart.TopRow = 7
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


