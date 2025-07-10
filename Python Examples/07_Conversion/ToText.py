from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ConversionSample2.xlsx"
outputFile = "ExceltoTxt.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet in excel workbook
sheet = workbook.Worksheets[0]
sheet.SaveToFile(outputFile, " ", Encoding.get_UTF8())
workbook.Dispose()

