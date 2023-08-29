from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/Template_Xls_3.xlsx"
outputFile = "UngroupExcelCells.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Ungroup the row 10 to 12.
sheet.UngroupByRows(10, 12)
#Ungroup the row 16 to 19.
sheet.UngroupByRows(16, 19)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

