from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/PictureBorder.xlsx"
outputFile = "RemovePictureBorder.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Get the first picture from the first worksheet
picture = sheet1.Pictures[0]
#Remove the picture border
#Method-1:
picture.Line.Visible = False
#Method-2:
#picture.Line.Weight = 0
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

