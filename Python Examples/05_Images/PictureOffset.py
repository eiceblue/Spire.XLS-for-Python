from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/logo.png"
outputFile = "PictureOffset.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
#Insert a picture
pic = sheet.Pictures.Add(2, 2,inputFile)
#Set left offset and top offset from the current range
pic.LeftColumnOffset = 200
pic.TopRowOffset = 100
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

