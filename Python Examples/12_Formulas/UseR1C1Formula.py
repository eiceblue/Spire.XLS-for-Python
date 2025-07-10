from spire.xls import *
from spire.xls.common import *


outputFile = "UseR1C1Formula.xlsx"

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
#Write array formula
sheet.Range["A5:C6"].FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)"
#Calculate Formulas
workbook.CalculateAllValue()
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
