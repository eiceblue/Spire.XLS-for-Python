from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_4.xlsx"
outputFile = "SetHeaderAndFooterMargins.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the PageSetup object of the first worksheet.
pageSetup = sheet.PageSetup
#Set the margins of header and footer.
pageSetup.HeaderMarginInch = 2
pageSetup.FooterMarginInch = 2
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()

