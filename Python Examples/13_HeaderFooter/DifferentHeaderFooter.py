from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/DifferentHeaderFooter.xlsx"
outputFile = "DifferentHeaderFooter.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#set text in range
sheet.Range["A1"].Text = "Page 1"
sheet.Range["G1"].Text = "Page 2"
#Set the different header footer for Odd and Even pages
sheet.PageSetup.DifferentOddEven = 1
#Set the header with font, size, bold and color
sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header"
sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer"
sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header"
sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer"
sheet.ViewMode = ViewMode.Layout
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

