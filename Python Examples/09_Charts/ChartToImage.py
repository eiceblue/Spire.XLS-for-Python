﻿from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartToImage.xlsx"
outputFile = "ChartToImage.png"

#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Save chart as image
image = workbook.SaveChartAsImage(workbook.Worksheets[0], 0)
image.Save(outputFile, ImageFormat.get_Png())
workbook.Dispose()
image.Dispose()


