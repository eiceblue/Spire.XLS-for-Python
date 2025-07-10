from spire.xls import *
from spire.xls.common import *


outputFile = "SubTotalFormula.xlsx"

#Create a workbook
workbook = Workbook()
#Get the first sheet
sheet = workbook.Worksheets[0]
sheet.Range["A1"].NumberValue = 1
sheet.Range["A2"].NumberValue = 2
sheet.Range["A3"].NumberValue = 3
sheet.Range["B1"].NumberValue = 4
sheet.Range["B2"].NumberValue = 5
sheet.Range["B3"].NumberValue = 6
sheet.Range["C1"].NumberValue = 7
sheet.Range["C2"].NumberValue = 8
sheet.Range["C3"].NumberValue = 9
#Add SUBTOTAL formulas
sheet.Range["A5"].Formula = "=SUBTOTAL(1,A1:C3)"
sheet.Range["B5"].Formula = "=SUBTOTAL(2,A1:C3)"
sheet.Range["C5"].Formula = "=SUBTOTAL(5,A1:C3)"
#Calculate Formulas
workbook.CalculateAllValue()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

