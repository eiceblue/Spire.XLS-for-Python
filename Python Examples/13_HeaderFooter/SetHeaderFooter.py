from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetHeaderFooter.xlsx"
outputFile = "SetHeaderFooter.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet
Worksheet = workbook.Worksheets[0]
#Set left header,"Arial Unicode MS" is font name, "18" is font size.
Worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .Python "
#Set center footer 
Worksheet.PageSetup.CenterFooter = "Footer Text"
Worksheet.ViewMode = ViewMode.Layout
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

