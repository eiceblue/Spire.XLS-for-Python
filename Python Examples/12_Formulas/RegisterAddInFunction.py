from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/Test.xlam"
outputFile = "RegisterAddInFunction.xlsx"

#Create a workbook
workbook = Workbook()
#Register AddIn function
workbook.AddInFunctions.Add(inputFile, "TEST_UDF")
workbook.AddInFunctions.Add(inputFile, "TEST_UDF1")
#Get the first sheet
sheet = workbook.Worksheets[0]
#Call AddIn function
sheet.Range["A1"].Formula = "=TEST_UDF()"
sheet.Range["A2"].Formula = "=TEST_UDF1()"
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

