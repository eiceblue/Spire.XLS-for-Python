from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChangeDataLabel.xlsx"
outputFile = "ChangeDataLabel.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Change data label of the frist datapoint of the first series
chart.Series[0].DataPoints[0].DataLabels.Text = "changed data label"
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

