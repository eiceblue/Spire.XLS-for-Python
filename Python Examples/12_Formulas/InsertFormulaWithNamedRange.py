from spire.xls import *
from spire.xls.common import *


outputFile = "InsertFormulaWithNamedRange.xlsx"

#Create a workbook
workbook = Workbook()
sheet = workbook.Worksheets[0]
#Set value
sheet.Range["A1"].Value = "1"
sheet.Range["A2"].Value = "1"
#Create a named range
NamedRange = workbook.NameRanges.Add("NewNamedRange")
NamedRange.NameLocal = "=SUM(A1+A2)"
#Set the formula
sheet.Range["C1"].Formula = "NewNamedRange"
#Save to file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

