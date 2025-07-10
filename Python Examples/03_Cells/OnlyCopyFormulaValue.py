from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CopyOnlyFormulaValue1.xlsx"
outputFile = "OnlyCopyFormulaValue.xlsx"


workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Set the copy option
copyOptions = CopyRangeOptions.OnlyCopyFormulaValue
sourceRange = sheet.Range["A6:E6"]
sheet.Copy(sourceRange, sheet.Range["A8:E8"], copyOptions)
sourceRange.Copy(sheet.Range["A10:E10"], copyOptions)
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

