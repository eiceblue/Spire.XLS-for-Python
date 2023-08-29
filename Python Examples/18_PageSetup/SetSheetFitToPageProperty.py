from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetSheetFitToPageProperty.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
pageSetup = sheet.PageSetup
#Set the FitToPagesTall property.
sheet.PageSetup.FitToPagesTall = 1
#Set the FitToPagesWide property.
sheet.PageSetup.FitToPagesWide = 1
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

