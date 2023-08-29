from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChangeSeriesColor.xlsx"
outputFile = "ChangeSeriesColor.xlsx"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first chart
chart = sheet.Charts[0]
#Get the second series
cs = chart.Series[1]
#Set the fill type
cs.Format.Fill.FillType = ShapeFillType.SolidColor
#Change the fill color
cs.Format.Fill.ForeColor = Color.get_Orange()
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

