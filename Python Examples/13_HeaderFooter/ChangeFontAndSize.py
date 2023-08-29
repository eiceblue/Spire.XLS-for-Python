from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChangeFontAndSizeForHeaderAndFooter.xlsx"
outputFile = "ChangeFontAndSize.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the new font and size for the header and footer
text = sheet.PageSetup.LeftHeader
#"Arial Unicode MS" is font name, "18" is font size
text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS "
sheet.PageSetup.LeftHeader = text
sheet.PageSetup.RightFooter = text
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

