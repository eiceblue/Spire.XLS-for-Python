from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartToImage.xlsx"
inputFile_Img = "./Demos/Data/SpireXls.png"
outputFile = "AddPictureInChart.xlsx"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Add the picture in chart
chart.Shapes.AddPicture(inputFile_Img)
#Save and launch result file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

