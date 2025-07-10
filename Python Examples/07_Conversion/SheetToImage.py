from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SheetToImage.xlsx"
outputFile = "SheetToImage.png"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save(outputFile)
workbook.Dispose()
