from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/DiscontinuousData.xlsx"
outputFile = "DiscontinuousData.xlsx"

#Load a Workbook from disk
book = Workbook()
book.LoadFromFile(inputFile)
#Get the first sheet
sheet = book.Worksheets[0]
#Add a chart
chart = sheet.Charts.Add(ExcelChartType.ColumnClustered)
chart.SeriesDataFromRange = False
#Set the position of chart
chart.LeftColumn = 1
chart.TopRow = 10
chart.RightColumn = 10
chart.BottomRow = 24
#Add a series
cs1 = chart.Series.Add()
#Set the name of the cs1
cs1.Name = sheet.Range["B1"].Value
#Set discontinuous values for cs1
cs1.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"])
cs1.Values = sheet.Range["B2:B3"].AddCombinedRange(sheet.Range["B5:B6"]).AddCombinedRange(sheet.Range["B8:B9"])
#Set the chart type
cs1.SerieType = ExcelChartType.ColumnClustered
#Add a series
cs2 = chart.Series.Add()
cs2.Name = sheet.Range["C1"].Value
cs2.CategoryLabels = sheet.Range["A2:A3"].AddCombinedRange(sheet.Range["A5:A6"]).AddCombinedRange(sheet.Range["A8:A9"])
cs2.Values = sheet.Range["C2:C3"].AddCombinedRange(sheet.Range["C5:C6"]).AddCombinedRange(sheet.Range["C8:C9"])
cs2.SerieType = ExcelChartType.ColumnClustered
chart.ChartTitle = "Chart"
chart.ChartTitleArea.Font.Size = 20
chart.ChartTitleArea.Color = Color.get_Black()
chart.PrimaryValueAxis.HasMajorGridLines = False
#Save and Launch
book.SaveToFile(outputFile, ExcelVersion.Version2010)
book.Dispose()

