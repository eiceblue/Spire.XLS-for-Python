from spire.xls import *
from spire.xls.common import *


outputFile = "ShowLeaderLine.xlsx"

#Create a workbook
workbook = Workbook()
workbook.Version = ExcelVersion.Version2013
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set value of specified range
sheet.Range["A1"].Value = "1"
sheet.Range["A2"].Value = "2"
sheet.Range["A3"].Value = "3"
sheet.Range["B1"].Value = "4"
sheet.Range["B2"].Value = "5"
sheet.Range["B3"].Value = "6"
sheet.Range["C1"].Value = "7"
sheet.Range["C2"].Value = "8"
sheet.Range["C3"].Value = "9"
chart = sheet.Charts.Add(ExcelChartType.BarStacked)
chart.DataRange = sheet.Range["A1:C3"]
chart.TopRow = 4
chart.LeftColumn = 2
chart.Width = 450
chart.Height = 300
for cs in chart.Series:
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
    cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
