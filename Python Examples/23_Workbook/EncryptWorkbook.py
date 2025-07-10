from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/EncryptWorkbook.xlsx"
outputFile = "EncryptWorkbook.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Protect Workbook with the password you want
workbook.Protect("eiceblue")
#Save the document and launch it
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
