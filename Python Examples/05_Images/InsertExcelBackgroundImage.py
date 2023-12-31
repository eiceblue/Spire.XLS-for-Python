﻿from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
inputFile_Img = "./Demos/Data/Background.png"
outputFile = "InsertExcelBackgroundImage.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Open an image. 
bm = Image.FromFile(inputFile_Img)
#Set the image to be background image of the worksheet.
sheet.PageSetup.BackgoundImage = bm
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()


