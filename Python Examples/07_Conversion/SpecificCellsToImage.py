from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ConversionSample1.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first wirksheet in Excel file
sheet = workbook.Worksheets[0]
#Specify Cell Ranges and Save to certain Image formats
sheet.ToImage(1, 1, 7, 5).Save( "SpecificCellsToImage.png")
sheet.ToImage(8, 1, 15, 5).Save( "SpecificCellsToImage.jpg")
sheet.ToImage(17, 1, 23, 5).Save( "SpecificCellsToImage.bmp")
