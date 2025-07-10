from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WriteHyperlinks.xlsx"
outputFile = "WriteHyperlinks.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
sheet.Range["B9"].Text = "Home page"
hylink1 = sheet.HyperLinks.Add(sheet.Range["B10"])
hylink1.Type = HyperLinkType.Url
hylink1.Address = """http://www.e-iceblue.com"""
sheet.Range["B11"].Text = "Support"
hylink2 = sheet.HyperLinks.Add(sheet.Range["B12"])
hylink2.Type = HyperLinkType.Url
hylink2.Address = "mailto:support@e-iceblue.com"
sheet.Range["B13"].Text = "Forum"
hylink3 = sheet.HyperLinks.Add(sheet.Range["B14"])
hylink3.Type = HyperLinkType.Url
hylink3.Address = "https://www.e-iceblue.com/forum/"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

