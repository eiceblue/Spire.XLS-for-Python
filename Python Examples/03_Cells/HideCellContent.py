from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_1.xlsx"
outputFile = "HideCellContent.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
#Hide the area by setting the number format as ";;;".
sheet.Range["C5:D6"].NumberFormat = ";;;"
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

