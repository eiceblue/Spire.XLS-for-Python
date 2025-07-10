from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Template_Xls_6.xlsx"
outputFile = "CreateFormulaConditionalFormat.xlsx"

#Create a workbook.
workbook = Workbook()
#Load the file from disk.
workbook.LoadFromFile(inputFile)
#Get the first worksheet and the first column from the workbook.
sheet = workbook.Worksheets[0]
range = sheet.Columns[0]
#Set the conditional formatting formula and apply the rule to the chosen cell range.
xcfs = sheet.ConditionalFormats.Add()
xcfs.AddRange(range)
conditional = xcfs.AddCondition()
conditional.FormatType = ConditionalFormatType.Formula
conditional.FirstFormula = "=($A1<$B1)"
conditional.BackKnownColor = ExcelColors.Yellow
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()

