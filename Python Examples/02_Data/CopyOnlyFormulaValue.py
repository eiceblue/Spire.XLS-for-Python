from spire.xls.common import *
from spire.xls import *



inputFile = "./Demos/Data/CopyOnlyFormulaValue.xlsx"
outputFile = "CopyOnlyFormulaValue.xlsx"
workbook = Workbook()
workbook.LoadFromFile(inputFile)

sheet = workbook.Worksheets[0]

#Set the copy option
copyOptions = CopyRangeOptions.OnlyCopyFormulaValue

#Copy range
sheet.Copy(sheet.Range["A2:C2"], sheet.Range["A5:C5"], copyOptions)
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()
