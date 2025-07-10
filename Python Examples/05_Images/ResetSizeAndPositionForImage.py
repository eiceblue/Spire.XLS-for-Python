from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SpireXls.png"
outputFile = "ResetSizeAndPositionForImage.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Add a picture to the first worksheet.
picture = sheet.Pictures.Add(1, 1, inputFile)
#Set the size for the picture.
picture.Width = 200
picture.Height = 200
#Set the position for the picture.
picture.Left = 200
picture.Top = 100
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

