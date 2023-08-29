from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample1.xlsx"
outputFile = "SetLegendBackgroundColor.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
ws = workbook.Worksheets[0]
chart = ws.Charts[0]
x = chart.Legend.FrameFormat if isinstance(chart.Legend.FrameFormat, XlsChartFrameFormat) else None
x.Fill.FillType = ShapeFillType.SolidColor
x.ForeGroundColor = Color.get_SkyBlue()
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

