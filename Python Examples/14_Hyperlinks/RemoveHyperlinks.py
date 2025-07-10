from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/HyperlinksSample1.xlsx"
outputFile = "RemoveHyperlinks.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Get the collection of all hyperlinks in the worksheet
links = sheet.HyperLinks
#Remove all link content
sheet.Range["B1"].ClearAll()
sheet.Range["B2"].ClearAll()
sheet.Range["B3"].ClearAll()
#Remove hyperlink and keep link text
sheet.HyperLinks.RemoveAt(0)
sheet.HyperLinks.RemoveAt(0)
sheet.HyperLinks.RemoveAt(0)
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
