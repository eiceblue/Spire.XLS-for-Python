from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetMarginsOfExcelSheet.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the PageSetup object of the first worksheet.
pageSetup = sheet.PageSetup
#Set bottom,left,right and top page margins.
pageSetup.BottomMargin = 2
pageSetup.LeftMargin = 1
pageSetup.RightMargin = 1
pageSetup.TopMargin = 3
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

