from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "SetImageOffsetOfChart.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
sheet1 = workbook.Worksheets.Add("Contrast")
#Add chart1 and background image to sheet1 as comparision.
chart1 = sheet1.Charts.Add(ExcelChartType.ColumnClustered)
chart1.DataRange = sheet.Range["D1:E8"]
chart1.SeriesDataFromRange = False
#Chart Position.
chart1.LeftColumn = 1
chart1.TopRow = 11
chart1.RightColumn = 8
chart1.BottomRow = 33
#Add picture as background.
chart1.ChartArea.Fill.CustomPicture(Stream("./Demos/Data/Background.png"), "None")
chart1.ChartArea.Fill.Tile = False
#Set the image offset.  
chart1.ChartArea.Fill.PicStretch.Left = 20
chart1.ChartArea.Fill.PicStretch.Top = 20
chart1.ChartArea.Fill.PicStretch.Right = 5
chart1.ChartArea.Fill.PicStretch.Bottom = 5
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


