from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SampeB_5.xlsx"
outputFile = "ChartAxisTitle.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Set axis title
chart.PrimaryCategoryAxis.Title = "Category Axis"
chart.PrimaryValueAxis.Title = "Value axis"
#Set font size
chart.PrimaryCategoryAxis.Font.Size = 12
chart.PrimaryValueAxis.Font.Size = 12
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

