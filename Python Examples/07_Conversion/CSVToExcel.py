from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/CSVToExcel.csv"
outputFile = "CSVToExcel.xlsx"

#Create a workbook
workbook = Workbook()
#Load a csv file
workbook.LoadFromFile(inputFile, ",", 1, 1)
sheet = workbook.Worksheets[0]
sheet.Range["D2:E19"].IgnoreErrorOptions = IgnoreErrorType.NumberAsText
sheet.AllocatedRange.AutoFitColumns()
#Save the document and launch it
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)


