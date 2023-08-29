from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "SetFormulaWithNamedRange.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the sheet
sheet = workbook.Worksheets[0]
#Create a named range
NamedRange = workbook.NameRanges.Add("MyNamedRange")
#Refers to range
NamedRange.RefersToRange = sheet.Range["B10:B12"]
#Set the formula of range to named range
sheet.Range["B13"].Formula = "=SUM(MyNamedRange)"
#Set value of ranges
sheet.Range["B10"].Value2 = Int32(10)
sheet.Range["B11"].Value2 = Int32(20)
sheet.Range["B12"].Value2 = Int32(30)
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

