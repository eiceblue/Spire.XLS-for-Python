from spire.xls import *
from spire.common import *


outputFile = "LinkToOtherSheetCell.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
range = sheet.Range["A1"]
#Add hyperlink in the range
hyperlink = sheet.HyperLinks.Add(range)
#Set the link type
hyperlink.Type = HyperLinkType.Workbook
#Set the display text
hyperlink.TextToDisplay = "Link to Sheet2 cell C5"
#Set the address
hyperlink.Address = "Sheet2!C5"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

