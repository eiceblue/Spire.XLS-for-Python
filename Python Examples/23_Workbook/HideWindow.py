from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/HideWindowExample.xlsx"
outputFile = "HideWindow.xlsx"

#Create a workbook
workbook = Workbook()
#Load an excel file
workbook.LoadFromFile(inputFile)
#Hide window
workbook.IsHideWindow = True
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

