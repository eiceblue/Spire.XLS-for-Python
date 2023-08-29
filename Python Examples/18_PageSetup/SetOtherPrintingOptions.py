from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "SetOtherPrintingOptions.xlsx"

#Create a workbook.
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Get the reference of the PageSetup of the worksheet.
pageSetup = sheet.PageSetup
#Allow to print gridlines.
pageSetup.IsPrintGridlines = True
#Allow to print row/column headings.
pageSetup.IsPrintHeadings = True
#Allow to print worksheet in black & white mode.
pageSetup.BlackAndWhite = True
#Allow to print comments as displayed on worksheet.
pageSetup.PrintComments = PrintCommentType.InPlace
#Allow to print worksheet with draft quality.
pageSetup.Draft = True
#Allow to print cell errors as N/A.
pageSetup.PrintErrors = PrintErrorsType.NA
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()


