from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "SetMargins.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set margins for top, bottom, left and right, here the unit of measure is Inch
sheet.PageSetup.TopMargin = 0.3
sheet.PageSetup.BottomMargin = 1
sheet.PageSetup.LeftMargin = 0.2
sheet.PageSetup.RightMargin = 1
#Set the header margin and footer margin
sheet.PageSetup.HeaderMarginInch = 0.1
sheet.PageSetup.FooterMarginInch = 0.5
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

