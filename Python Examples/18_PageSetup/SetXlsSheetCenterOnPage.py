from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetXlsSheetCenterOnPage.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the PageSetup object of the first page.
pageSetup = sheet.PageSetup
#Set the worksheet center on page.
pageSetup.CenterHorizontally = True
pageSetup.CenterVertically = True
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

