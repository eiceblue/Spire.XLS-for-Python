from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetPrintAreaOfXlsFile.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Specify the cells range of the print area.
pageSetup.PrintArea = "A1:E5"
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()


