from spire.xls.common import *
from spire.xls import *


inputFile = "./Demos/Data/AddATotalRowToTable.xlsx"
outputFile = "AddTotalRowToTable.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Create a table with the data from the specific cell range.
table = sheet.ListObjects.Create("Table", sheet.Range["A1:D4"])
#Display total row.
table.DisplayTotalRow = True
#Add a total row.
cols =table.Columns
cols[0].TotalsRowLabel = "Total"
cols[1].TotalsCalculation = ExcelTotalsCalculation.Sum
cols[2].TotalsCalculation = ExcelTotalsCalculation.Sum
cols[3].TotalsCalculation = ExcelTotalsCalculation.Sum
#Save to file.
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()