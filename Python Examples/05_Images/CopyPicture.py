from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "CopyPicture.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet1 = workbook.Worksheets[0]
#Add a new worksheet as destination sheet
destinationSheet = workbook.Worksheets.Add("DestSheet")
#Get the first picture from the first worksheet
sourcePicture = sheet1.Pictures[0]
#Get the image
image = sourcePicture.Picture
#Add the image into the added worksheet 
destinationSheet.Pictures.Add(2, 2, image)
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

