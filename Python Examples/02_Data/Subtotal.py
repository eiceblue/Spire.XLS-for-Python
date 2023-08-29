from spire.xls import *
from spire.common import *

inputFile = "./Demos/Data/Subtotal.xlsx"
outputFile = "Subtotal.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Select data range
range = sheet.Range["A1:B18"]
#Subtotal selected data
sheet.Subtotal(range, 0, [1], SubtotalTypes.Sum, True, False, True)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

