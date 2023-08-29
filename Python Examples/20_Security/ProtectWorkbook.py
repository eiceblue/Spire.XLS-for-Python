from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/WorksheetSample1.xlsx"
outputFile = "ProtectWorkbook.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Protect Workbook
workbook.Protect("e-iceblue")
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

