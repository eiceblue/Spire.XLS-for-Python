from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CommonTemplate1.xlsx"
outputFile = "AddHyperlinkToText.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Add url link
UrlLink = sheet.HyperLinks.Add(sheet.Range["D10"])
UrlLink.TextToDisplay = sheet.Range["D10"].Text
UrlLink.Type = HyperLinkType.Url
UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago"
#Add email link
MailLink = sheet.HyperLinks.Add(sheet.Range["E10"])
MailLink.TextToDisplay = sheet.Range["E10"].Text
MailLink.Type = HyperLinkType.Url
MailLink.Address = "mailto:Amor.Aqua@gmail.com"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

