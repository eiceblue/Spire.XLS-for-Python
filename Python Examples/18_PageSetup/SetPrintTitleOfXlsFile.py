from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetPrintTitleOfXlsFile.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
pageSetup = sheet.PageSetup
#Define column numbers A & B as title columns.
pageSetup.PrintTitleColumns = "$A:$B"
#Defining row numbers 1 & 2 as title rows.
pageSetup.PrintTitleRows = "$1:$2"
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

