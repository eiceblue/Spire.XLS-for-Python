from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/LocateImages.xlsx"
outputFile = "LocateImages.xlsx"

#Create a Workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
pic = sheet.Pictures[0]
pic.LeftColumnOffset = 300
pic.TopRowOffset = 300
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
