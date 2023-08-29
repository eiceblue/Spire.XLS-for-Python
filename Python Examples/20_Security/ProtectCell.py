from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ProtectCell.xlsx"
outputFile = "ProtectCell.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Protect cell
sheet.Range["B3"].Style.Locked = True
sheet.Range["C3"].Style.Locked = False
sheet.Protect("TestPassword", SheetProtectionType.All)
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
