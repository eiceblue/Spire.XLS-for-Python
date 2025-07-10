from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "CopyCellFormat.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Copy the cell format from column 2 and apply to cells of column 5.
count = sheet.Rows.Length
i = 1
while i < count + 1:
    sheet.Range["E{0}".format(i)].Style = sheet.Range["B{0}".format(i)].Style
    i += 1
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

