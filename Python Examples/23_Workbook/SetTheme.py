from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/SetTheme.xlsx"
outputFile = "SetTheme.xlsx"

#Create a workbook
srcWorkbook = Workbook()
#Load an excel file
srcWorkbook.LoadFromFile(inputFile)
srcWorksheet = srcWorkbook.Worksheets[0]
workbook = Workbook()
workbook.Worksheets.Clear()
workbook.Worksheets.AddCopy(srcWorksheet)
#1. Copy the theme of the workbook
#workbook.CopyTheme(srcWorkbook)
#2. Set a certain type of color of the default theme in the workbook
workbook.SetThemeColor(ThemeColorType.Dk1, Color.get_SkyBlue())
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

