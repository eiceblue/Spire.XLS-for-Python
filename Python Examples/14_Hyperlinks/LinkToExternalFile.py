from spire.xls import *
from spire.xls.common import *

inputFile ="./Demos/Data/SampeB_4.xlsx"
outputFile = "LinkToExternalFile.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
range = sheet.Range[1,1]
#Add hyperlink in the range
hyperlink = sheet.HyperLinks.Add(range)
#Set the link type
hyperlink.Type = HyperLinkType.File
#Set the display text
hyperlink.TextToDisplay = "Link To External File"
#Set file address
hyperlink.Address = inputFile
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

