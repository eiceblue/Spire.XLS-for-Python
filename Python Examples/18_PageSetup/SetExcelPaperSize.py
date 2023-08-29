from spire.xls import *
from spire.common import *


outputFile = "SetExcelPaperSize.xlsx"

#Create a workbook.
workbook = Workbook()
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Set the paper size of the worksheet as A4 paper.
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
workbook.SaveToFile(outputFile,ExcelVersion.Version2010)
workbook.Dispose()


