from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetPrintQualityOfXlsFile.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the print quality of the worksheet to 180 dpi.
sheet.PageSetup.PrintQuality = 180
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

