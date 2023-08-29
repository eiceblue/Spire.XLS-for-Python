from spire.xls import *
from spire.common import *

inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "ReadImages.jpg"

#Create a Workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the first image
pic = sheet.Pictures[0]
#save
pic.Picture.Save(outputFile, ImageFormat.get_Jpeg())
workbook.Dispose()

