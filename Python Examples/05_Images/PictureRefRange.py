from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/PictureRefRange.xlsx"
outputFile = "PictureRefRange.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
sheet.Range["A1"].Value = "Spire.XLS"
sheet.Range["B3"].Value = "E-iceblue"
#Get the first picture in worksheet
picture = sheet.Pictures[0]
#Set the reference range of the picture to A1:B3
picture.RefRange = "A1:B3"
#Save and Launch
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

