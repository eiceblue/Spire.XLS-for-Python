from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/SetDataValidationOnSeparateSheet.xlsx"
outputFile = "SetDataValidationOnSeparateSheet.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#This is the first sheet
sheet1 = workbook.Worksheets[0]
sheet1.Range["B10"].Text = "Here is a dataValidation example."
#This is the second sheet
sheet2 = workbook.Worksheets[1]
#The property is to enable the data can be from different sheet.
sheet2.ParentWorkbook.Allow3DRangesInDataValidation = True
sheet1.Range["B11"].DataValidation.DataRange = sheet2.Range["A1:A7"]
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
