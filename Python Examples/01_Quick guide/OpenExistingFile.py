from spire.common import *
from spire.xls import *


inputFile = "./Demos/Data/templateAz2.xlsx"
outputFile = "OpenExistingFile.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets.Add("MySheet")
sheet.Range["A1"].Text = "Hello World"
workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()