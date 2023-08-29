from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/UnprotectProtectSheet.xlsx"
outputFile = "UnlockProtectSheet.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Unlock the worksheet in a unlocked Excel file with null string.
sheet.Unprotect("e-iceblue")
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

