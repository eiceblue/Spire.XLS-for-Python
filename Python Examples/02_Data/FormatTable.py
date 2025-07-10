from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/FormatTable.xlsx"
outputFile = "FormatTable.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Add Default Style to the table
sheet.ListObjects[0].BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9
#Show Total
sheet.ListObjects[0].DisplayTotalRow = True
#Set calculation type
sheet.ListObjects[0].Columns[0].TotalsRowLabel = "Total"
sheet.ListObjects[0].Columns[1].TotalsCalculation = ExcelTotalsCalculation.none
sheet.ListObjects[0].Columns[2].TotalsCalculation = ExcelTotalsCalculation.none
sheet.ListObjects[0].Columns[3].TotalsCalculation = ExcelTotalsCalculation.Sum
sheet.ListObjects[0].Columns[4].TotalsCalculation = ExcelTotalsCalculation.Sum

sheet.ListObjects[0].ShowTableStyleRowStripes = True

sheet.ListObjects[0].ShowTableStyleColumnStripes = True
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

