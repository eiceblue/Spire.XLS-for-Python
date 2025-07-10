from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CompressPictures.xlsx"
outputFile = "CompressPictures.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
for sheet in workbook.Worksheets:
    for picture in sheet.Pictures:
        picture.Compress(50)
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
