from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SampleB_2.xlsx"
outputFile = "ToImageWithoutWhiteSpace.png"

#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Set the margin as 0 to remove the white space around the image
sheet.PageSetup.LeftMargin = 0
sheet.PageSetup.BottomMargin = 0
sheet.PageSetup.TopMargin = 0
sheet.PageSetup.RightMargin = 0
#convert to image
image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)
image.Save(outputFile)


