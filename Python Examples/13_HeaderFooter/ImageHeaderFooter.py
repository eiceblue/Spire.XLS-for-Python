from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ImageHeaderFooter.xlsx"
inputImage = "./Demos/Data/Logo.png"
outputFile = "ImageHeaderFooter.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Load an image from disk
image = Stream(inputImage)
#Set the image header
sheet.PageSetup.LeftHeaderImage = image
sheet.PageSetup.LeftHeader = "&G"
#Set the image footer
sheet.PageSetup.CenterFooterImage = image
sheet.PageSetup.CenterFooter = "&G"
#Set the view mode of the sheet
sheet.ViewMode = ViewMode.Layout
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
