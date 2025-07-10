from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/DecryptWorkbook.xlsx"
outputFile = "DecryptWorkbook.xlsx"

# Detect if the Excel workbook is password protected.
outValue = Workbook.IsPasswordProtected(inputFile)

if outValue:
    # Load a file with the password specified
    workbook =  Workbook()
    workbook.OpenPassword ="eiceblue"
    workbook.LoadFromFile(inputFile)

    # Decrypt workbook
    workbook.UnProtect()
    workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
