from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "OutputSetExcelPageOrderType.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Set the order type of the pages to over then down.
pageSetup.Order = OrderType.OverThenDown
workbook.SaveToFile(outputFile,ExcelVersion.Version2013)
workbook.Dispose()

