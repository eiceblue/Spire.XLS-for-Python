from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "ChangeMajorGridlines.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Change the color of marjor gridlines
chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.get_Red()
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

