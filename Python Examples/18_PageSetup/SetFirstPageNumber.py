from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetFirstPageNumber.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the first page number of the worksheet pages.
sheet.PageSetup.FirstPageNumber = 2
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

