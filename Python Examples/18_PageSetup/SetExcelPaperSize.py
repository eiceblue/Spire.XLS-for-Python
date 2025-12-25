from spire.xls.common import *
from spire.xls import *


outputFile = "SetExcelPaperSize.xlsx"

#Create a workbook.
workbook = Workbook()

#Get the first worksheet.
sheet = workbook.Worksheets[0]

#Set the paper size for the worksheet(e.g., PaperA0, PaperA1, PaperA3, PaperA4, etc.).
sheet.PageSetup.PaperSize = PaperSizeType.PaperA4

#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
