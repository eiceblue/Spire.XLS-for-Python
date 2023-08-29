from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WriteImages.xlsx"
inputFile_Img = "./Demos/Data/SpireXls.png"
outputFile = "WriteImages.xlsx"

#Create a Workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Add an image to the specific cell
sheet.Pictures.Add(14, 5, inputFile_Img)
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()


